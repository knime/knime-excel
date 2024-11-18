/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME AG, Zurich, Switzerland
 *  Website: http://www.knime.com; Email: contact@knime.com
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License, Version 3, as
 *  published by the Free Software Foundation.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program; if not, see <http://www.gnu.org/licenses>.
 *
 *  Additional permission under GNU GPL version 3 section 7:
 *
 *  KNIME interoperates with ECLIPSE solely via ECLIPSE's plug-in APIs.
 *  Hence, KNIME and ECLIPSE are both independent programs and are not
 *  derived from each other. Should, however, the interpretation of the
 *  GNU GPL Version 3 ("License") under any applicable laws result in
 *  KNIME and ECLIPSE being a combined program, KNIME AG herewith grants
 *  you the additional permission to use and propagate KNIME together with
 *  ECLIPSE with only the license terms in place for ECLIPSE applying to
 *  ECLIPSE and the GNU GPL Version 3 applying for KNIME, provided the
 *  license terms of ECLIPSE themselves allow for the respective use and
 *  propagation of ECLIPSE together with KNIME.
 *
 *  Additional permission relating to nodes for KNIME that extend the Node
 *  Extension (and in particular that are based on subclasses of NodeModel,
 *  NodeDialog, and NodeView) and that only interoperate with KNIME through
 *  standard APIs ("Nodes"):
 *  Nodes are deemed to be separate and independent programs and to not be
 *  covered works.  Notwithstanding anything to the contrary in the
 *  License, the License does not apply to Nodes, you are not required to
 *  license Nodes under the License, and you are granted a license to
 *  prepare and propagate Nodes, in each case even if such Nodes are
 *  propagated with or for interoperation with KNIME.  The owner of a Node
 *  may freely choose the license terms applicable to such Node, including
 *  when such Node is propagated with or for interoperation with KNIME.
 * ---------------------------------------------------------------------
 *
 * History
 *   Oct 20, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CancellationException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.message.Message;
import org.knime.core.util.FileUtil;
import org.knime.core.util.ThreadUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.connections.meta.FSDescriptorRegistry;
import org.knime.filehandling.core.connections.uriexport.URIExporterID;
import org.knime.filehandling.core.connections.uriexport.URIExporterIDs;
import org.knime.filehandling.core.connections.uriexport.noconfig.EmptyURIExporterConfig;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;
import org.knime.filehandling.core.node.table.reader.read.Read;

/**
 * Abstract class for {@link Read}s that read Excel files/spreadsheets.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class ExcelRead implements Read<ExcelCell> {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelRead.class);

    private static final int BLOCKING_QUEUE_SIZE = 100;

    private static final AtomicLong CACHED_THREAD_POOL_INDEX = new AtomicLong();

    /** Threadpool that can be used to parse new sheets. */
    private static final ExecutorService CACHED_THREAD_POOL = Executors.newCachedThreadPool(
        r -> ThreadUtils.threadWithContext(r, "KNIME-Excel-Parser-" + CACHED_THREAD_POOL_INDEX.getAndIncrement()));

    /** The poison pill put into the blocking to indicate end of parsing. */
    static final RandomAccessible<ExcelCell> POISON_PILL =
        RandomAccessibleUtils.createFromArray(new ExcelCell(KNIMECellType.STRING, "POISON"));

    /** Queue that uses the TRF to consume rows produced by the parser thread. */
    private final BlockingQueue<RandomAccessible<ExcelCell>> m_queueRandomAccessibles =
        new ArrayBlockingQueue<>(BLOCKING_QUEUE_SIZE);

    /** The thread running the parser. */
    private Future<?> m_parserThread;

    /** Atomic reference used to communicate exceptions between parser thread and main thread. */
    private final AtomicReference<Throwable> m_throwableDuringParsing = new AtomicReference<>();

    /** Iterator collecting RandomAccessibles from the parser thread. */
    private final RandomAccessibleIterator m_randomAccessibleIterator = new RandomAccessibleIterator();

    /** The path of the underlying source. */
    private final Path m_path;

    /** The Excel table read config. */
    protected final TableReadConfig<ExcelTableReaderConfig> m_config;

    /** The (local) file to read from. */
    // it is a future so the download task can be canceled with the same mechanism as the parser
    private Future<File> m_localFile;

    protected final Consumer<Map<String, Boolean>> m_sheetNamesConsumer;

    /**
     * Constructor
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @param sheetNamesConsumer
     * @throws IOException if a stream can not be created from the provided file.
     */
    protected ExcelRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
        final Consumer<Map<String, Boolean>> sheetNamesConsumer) throws IOException {
        m_path = path;
        m_config = config;

        if (path instanceof FSPath fsPath) {
            // our code always calls the ExcelRead with FSPath instances
            m_localFile = resolveToLocalOrTempFile(fsPath);
        } else {
            // in the rare case that we get a non-FSPath, we just assume that the file is already local
            m_localFile = CompletableFuture.completedFuture(path.toFile());
        }
        m_sheetNamesConsumer = sheetNamesConsumer;

        startParserThread();
    }

    /**
     * Resolves the given path to a local file if possible or creates a temporary copy of the referenced file.
     * The temporary file is deleted when the current file system is closed.
     *
     * @return future for local file or temporary local file with contents of path
     * @throws IOException I/O exception while resolving local file, creating temporary file, or transferring contents
     */
    /* AP-20714: We need to cache the file locally, since the InputStream-based API of OPCPackage puts the whole
         decompressed contents into main memory, which can exceed the heap space (and was observed to do for
         some files with ~80-300 MiB on disk).
         Using the file-based API, the OPCPackage is able to do random I/O on the contents of the compressed
         file.
         What did not work:
         - Using something like SharedStrings to deduplicate strings in memory is not possible,
           since this applies _after_ opening the OPCPackage -- and opening is the problem.
         - POI 5.1.0 adds some possibility to cache package parts (metadata, sheets, etc.) in (encrypted) temp
           files if they are above a certain threshold by setting
           `ZipInputStreamZipEntrySource#setThresholdBytesForTempFiles` and `ZipPackage#setUseTempFilePackageParts`:
           - (default) -1: "temp files are not used and that zip entries with more than 2GB of data after
                            decompressing will fail"
           - 0: "all zip entries are stored in temp files"
           When using the InputStream-based API, the contents get wrapped in "zip archive fake entries".
           For caching in temp files, this asks its wrapped ZipArchiveEntry for its size, which reports a size
           of -1 (SIZE_UNKNOWN) until it was decompressed fully. From the ZipArchiveEntry Javadoc (commons-compress):
           "Note: ZipArchiveInputStream may create entries that return SIZE_UNKNOWN as long as the entry hasn't been
            read completely."
           This is the case in the constructor of `ZipArchiveFakeEntry`.

           Hence, the reported size of -1 is not above any threshold that one can set to enable caching (anything >= 0).
           (see https://github.com/apache/poi/blob/trunk/poi-ooxml/src/main/java/org/apache/poi/openxml4j/util/ZipArchiveFakeEntry.java#L60) // NOSONAR

           But even if using InputStream would work, we would unnecessarily cache local files in temp files just
           because we have to go through InputStream.
           (for configuration options of POI see https://poi.apache.org/components/configuration.html)
     */
    private static Future<File> resolveToLocalOrTempFile(final FSPath path) {
        if (!Files.isRegularFile(path)) {
            throw new IllegalArgumentException(
                "Can only resolve regular files, not path denoting: \"%s\"".formatted(path.toString()));
        }
        return resolveToLocal(path)
                .<Future<File>>map(CompletableFuture::completedFuture)
                .orElseGet(() ->
                    // package the download (copy to temp file) into a callable to make it cancelable through the UI
                    CACHED_THREAD_POOL.submit(ThreadUtils.callableWithContext(() -> copyToTemp(path))
                ));
    }

    /**
     * Tries to resolve the given path to a file on the local machine.
     *
     * @param path path to a file
     * @return file on the local machine or {@link Optional#empty()} if the file could not be resolved to a local file
     */
    private static Optional<File> resolveToLocal(final FSPath path) {
        try {
            @SuppressWarnings("resource") // don't close the FS since we're not the ones who opened it
            final var desc = FSDescriptorRegistry.getFSDescriptor(path.getFileSystem().getFSType()).orElseThrow();
            for (final var exporterId : List.of(URIExporterIDs.KNIME_FILE, URIExporterIDs.LEGACY_KNIME_URL,
                    new URIExporterID("knime-customurl"))) {
                final var factory = desc.getURIExporterFactory(exporterId);
                if (factory != null) {
                    // we can determine whether a file is local from the first two exporters, so we can terminate early
                    final var exporter = factory.createExporter(EmptyURIExporterConfig.getInstance());
                    final var exportedUrl = exporter.toUri(path).toURL();
                    return Optional.ofNullable(FileUtil.getFileFromURL(exportedUrl));
                }
            }
        } catch (final URISyntaxException | MalformedURLException | IllegalArgumentException e) {
            LOGGER.debug("Problem resolving FSPath to local file, assuming non-local", e);
        }
        return Optional.empty();
    }

    /**
     * Copies the given file to a temp file using {@link Files#copy(Path, Path, java.nio.file.CopyOption...)}.
     * The temp file is deleted when the file system of the given path is closed (or the JVM exits).
     *
     * @param path path to copy
     * @return temp file
     * @throws IOException while copying
     * @see {@link Files#copy(Path, Path, java.nio.file.CopyOption...)}
     */
    private static File copyToTemp(final FSPath path) throws IOException {
        final var fileName = path.getFileName().toString();
        // we use the real extension so any file type detection that uses the extension will work
        final var ext = FilenameUtils.getExtension(fileName);
        final var tempFile =
            FileUtil.createTempFile("KNIMEExcelReaderRemoteTempCopy", "." + ext, FileUtil.getWorkflowTempDir(), true);
        final var tempPath = tempFile.toPath();
        try (final var fileSystem = path.getFileSystem()) {
            fileSystem.registerCloseable(() -> Files.deleteIfExists(tempPath));
        }
        LOGGER.debug(() -> "Caching Excel file at \"%s\" to temporary file \"%s\"".formatted(path, tempFile));
        Files.copy(path, tempPath, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.COPY_ATTRIBUTES);
        return tempFile;
    }

    /**
     * Creates a {@code ExcelParserRunnable} on a local file.
     *
     * @param localFile the file to read from
     * @return the {@code ExcelParserRunnable}
     * @throws IOException if an I/O exception occurs
     *
     * @see OPCPackage#open(File)
     * @see OPCPackage#open(java.io.InputStream)
     */
    protected abstract ExcelParserRunnable createParser(final File localFile) throws IOException;

    @Override
    public final RandomAccessible<ExcelCell> next() throws IOException {
        // check and return next element
        final boolean hasNext = m_randomAccessibleIterator.hasNext();
        if (m_throwableDuringParsing.get() != null) {
            close();
        }
        return hasNext ? m_randomAccessibleIterator.next() : null;
    }

    /**
     * Creates and returns an {@link IOException} with an error message telling the user which file requires a password
     * to be opened.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    protected IOException createPasswordProtectedFileException(final EncryptedDocumentException e) {
        return new IOException(String
            .format("The file '%s' is password protected. Supply a password via the encryption settings.", m_path), e);
    }

    /**
     * Creates and returns an {@link IOException} with an error message telling the user for which file the password is
     * incorrect.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    protected IOException createPasswordIncorrectException(final EncryptedDocumentException e) {
        return new IOException(String.format("The supplied password is incorrect for file '%s'.", m_path), e);
    }

    private void startParserThread() throws IOException {
        try {
            // create and start the thread
            final ExcelParserRunnable runnable = createParser(getFile());
            m_parserThread = CACHED_THREAD_POOL.submit(ThreadUtils.runnableWithContext(runnable));
        } catch (InvalidOperationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private File getFile() throws IOException {
        try {
            return m_localFile.get();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("Waiting for local file to be available was interrupted: " + e.getMessage(), e);
        } catch (ExecutionException e) {
            throw new IOException("Task supplying the local file was aborted by exception: " + e.getMessage(), e);
        }
    }

    @Override
    public final void close() throws IOException {
        // cancel the thread
        if (m_parserThread != null && m_parserThread.cancel(true)) {
            LOGGER.debug("Canceled parser thread");
            try {
                // wait for the parser thread to cancel, to avoid it still reading from the (sheet) input stream
                // that we will make sure is closed as well (in closeResources)
                m_parserThread.get();
            } catch (final InterruptedException e) {
                Thread.currentThread().interrupt();
            } catch (final ExecutionException e) {
                LOGGER.debug("While waiting for parser thread to cancel", e);
            } catch (final CancellationException e) { // NOSONAR
                // ignore since we wanted to cancel it
            }
        }

        final var throwable = m_throwableDuringParsing.get();
        // if an exception occurred during parsing, throw it
        if (throwable != null) {
            if (throwable instanceof CorruptExcelFileException cex) {
                throw Message.builder() //
                    .addResolutions("Try to repair the file using Excel") //
                    .withSummary("The Excel file being read is corrupt and cannot be read: " + cex.getMessage()) //
                    .build().orElseThrow() //
                    .toKNIMEException(cex).toUnchecked();
            } else if (throwable instanceof RuntimeException rex) {
                throw rex;
            } else if (throwable instanceof Error err) {
                throw err;
            }
            throw new IllegalStateException(throwable.getMessage(), throwable);
        }
    }

    /**
     * Adds the {@link RandomAccessible} to the blocking queue.
     *
     * @param randomAccessible the {@link RandomAccessible} to add
     * @throws InterruptedException if {@link ArrayBlockingQueue#put(Object)} is interrupted while waiting
     */
    protected void addToQueue(final RandomAccessible<ExcelCell> randomAccessible) throws InterruptedException {
        m_queueRandomAccessibles.put(randomAccessible);
    }

    /**
     * Sets the throwable.
     *
     * @param throwable the {@link Throwable} to set
     */
    protected void setThrowable(final Throwable throwable) {
        m_throwableDuringParsing.set(throwable);
    }

    /**
     * Iterator that collects all parsed rows. If an exception occurred during parsing or the parsing is finished,
     * {@link #hasNext()} will return {@code false}. Otherwise, it will wait for more rows becoming available to iterate
     * over.
     */
    private class RandomAccessibleIterator implements Iterator<RandomAccessible<ExcelCell>> {

        private final LinkedList<RandomAccessible<ExcelCell>> m_randomAccessibles = new LinkedList<>();

        private boolean m_encounteredPoisonPill;

        @Override
        public boolean hasNext() {
            if (m_encounteredPoisonPill || m_throwableDuringParsing.get() != null) {
                return false;
            }
            // TODO in my experiments, the size of the queue was never bigger than 1 -> wouldn't a simple #take be
            // faster? -> check that once we have more features implemented with different xlsx files
            if (m_randomAccessibles.isEmpty()) {
                try {
                    m_randomAccessibles.add(m_queueRandomAccessibles.take());
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    return false;
                }
                // add all elements of the queue into the iterator's list
                m_queueRandomAccessibles.drainTo(m_randomAccessibles);
            }
            // poison pill indicates end of parsing (normall or otherwise)
            m_encounteredPoisonPill = m_randomAccessibles.peek() == POISON_PILL;
            return !m_encounteredPoisonPill;
        }

        @Override
        public RandomAccessible<ExcelCell> next() {
            if (!hasNext()) {
                throw new NoSuchElementException("No more rows available.");
            }
            return m_randomAccessibles.poll();
        }

    }

    /**
     * Returns the name of the selected sheet based on the configuration.
     *
     * @return the name of the selected sheet
     * @throws IOException if the configured sheet is not available
     */
    protected String getSelectedSheet(final Map<String, Boolean> sheetNames) throws IOException {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        return excelConfig.getSheetSelection().getSelectedSheet(sheetNames, excelConfig, m_path);
    }

    /**
     * Returns the indexes of the hidden columns.
     *
     * @return the indexes of the hidden columns in a {@link Set}
     */
    public abstract Set<Integer> getHiddenColumns();

}
