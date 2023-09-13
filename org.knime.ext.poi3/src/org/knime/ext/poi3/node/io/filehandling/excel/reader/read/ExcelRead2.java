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
 *   12 Sep 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.OptionalLong;
import java.util.Set;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.LongConsumer;

import org.apache.poi.EncryptedDocumentException;
import org.knime.core.node.NodeLogger;
import org.knime.core.util.ThreadUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;
import org.knime.filehandling.core.node.table.reader.read.Read;

/**
 * {@link Read}s that read Excel files/spreadsheets using a companion thread for the parser.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelRead2 implements Read<ExcelCell> {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelRead2.class);

    private static final int BLOCKING_QUEUE_SIZE = 100;

    private static final AtomicLong CACHED_THREAD_POOL_INDEX = new AtomicLong();

    /** The poison pill put into the blocking to indicate end of parsing. */
    static final RandomAccessible<ExcelCell> POISON_PILL =
        RandomAccessibleUtils.createFromArray(new ExcelCell(KNIMECellType.STRING, "POISON"));

    /** Queue that uses the TRF to consume rows produced by the parser thread. */

    private final BlockingQueue<RandomAccessible<ExcelCell>> m_queueRandomAccessibles =
            new ArrayBlockingQueue<>(BLOCKING_QUEUE_SIZE);

    /* Why a companion thread?
     *
     * - XLS:
     *   The only way to read thos is to open the workbook, which means we could read them iterator/pull-based, not
     *   needing a separate thread.
     *
     * - XLSX:
     *   POI's solution to read XLSX in a memory-efficient way is to delegate to SAX or StaX (very low level).
     *   For SAX it already offers event handlers that do almost everything that we need.
     *
     *   SAX is event/push-based, whereas StaX is iterator/pull-based. Since we currently use the former, we need the
     *   companion thread to host the parser.
     *   Technically, we could read sheets from XLSX files by opening it as ZIP and reading the sheet's XML file with
     *   a SAX/StaX parser. So there is currently no technical reason to use SAX other than that the implementation is
     *   already there. If we wanted to remove the companion thread, we would need to refactor the event-based handlers
     *   into a StaX parser.
     *
     * - XLSB:
     *   POI currently only offers an event-based handler for binary Excel files and nothing iterative.
     *   Therefore, in order to move away from the companion thread we would need to adopt the handler and transform
     *   it into a iterator-based parser.
    **/
    /** The thread running the parser. */
    private Thread m_parserThread;

    private final AtomicReference<Throwable> m_throwableDuringParsing = new AtomicReference<>();

    /** Iterator collecting RandomAccessibles from the parser thread. */
    private final RandomAccessibleIterator m_randomAccessibleIterator = new RandomAccessibleIterator();

    /** The path of the underlying source. */
    private final Path m_path;

    /** The Excel table read config. */
    private final TableReadConfig<ExcelTableReaderConfig> m_config;


    private AtomicLong m_currentProgress = new AtomicLong();

    private CompletableFuture<WorkbookMetadata> m_metadata = new CompletableFuture<>();

    public record SheetMetadata(long sizeInBytes, Set<Integer> hiddenColumns) {}

    /**
     * Record for holding metadata about the workbook and current sheet
     */
    public record WorkbookMetadata(Map<String, Boolean> sheetNames, SheetMetadata sheetMeta) {}


    /**
     * Constructor
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     */
    public ExcelRead2(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) {
        m_path = path;
        m_config = config;
        startParserThread();
    }

    @Override
    public RandomAccessible<ExcelCell> next() throws IOException {
        final boolean hasNext = m_randomAccessibleIterator.hasNext();
        if (!m_parserThread.isAlive()) {
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

    private void startParserThread() {
        m_parserThread = ThreadUtils.threadWithContext(createParser(m_throwableDuringParsing::set, m_metadata),
            "KNIME-Excel-Parser-" + CACHED_THREAD_POOL_INDEX.getAndIncrement());
        m_parserThread.start();
    }


    public class CellConsumer {

        private final LongConsumer m_progress;

        public CellConsumer(final LongConsumer progress) {
            m_progress = progress;
        }

        void finished() throws InterruptedException {
            addToQueue(POISON_PILL);
        }

        public void accept(final RandomAccessible<ExcelCell> row, final long progress) throws InterruptedException {
            addToQueue(row);
            if (m_progress != null) {
                m_progress.accept(progress);
            }
        }
    }

    /**
     * @param metadata
     * @param path
     * @return
     */
    private ExcelParserRunnable2 createParser(final Consumer<Throwable> exceptionHandler,
            final CompletableFuture<WorkbookMetadata> metadata) {
        return new ExcelParserRunnable2(m_path, m_config, new CellConsumer(m_currentProgress::set), metadata,
            exceptionHandler);
    }

    @Override
    public void close() throws IOException {
        // make sure the thread is finished so it has freed its resources
        m_parserThread.interrupt();
        InterruptedException ownException = null;
        try {
            m_parserThread.join();
        } catch (InterruptedException e) {
            ownException = e;
        }

        // if an exception occurred during parsing, throw it preferably
        final var throwable = m_throwableDuringParsing.get();
        if (throwable != null) {
            if (ownException != null) {
                throwable.addSuppressed(ownException);
            }
            if (throwable instanceof RuntimeException rex) {
                throw rex;
            } else if (throwable instanceof Error err) {
                throw err;
            }
            final var msg = throwable.getMessage();
            if (msg != null) {
                throw new IOException(msg, throwable);
            }
            throw new IOException(throwable);
        } else {
            if (ownException != null) {
                throw new IOException(ownException);
            }
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
     * Iterator that collects all parsed rows. If an exception occurred during parsing or the parsing is finished,
     * {@link #hasNext()} will return {@code false}. Otherwise, it will wait for more rows becoming available to iterate
     * over.
     */
    private class RandomAccessibleIterator implements Iterator<RandomAccessible<ExcelCell>> {

        private final LinkedList<RandomAccessible<ExcelCell>> m_randomAccessibles = new LinkedList<>();

        private boolean m_encounteredPoisonPill;

        @Override
        public boolean hasNext() {
            if (m_encounteredPoisonPill) {
                return false;
            }
            // TODO in my experiments, the size of the queue was never bigger than 1 -> wouldn't a simple #take be faster?
            // -> check that once we have more features implemented with different xlsx files
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

    @Override
    public OptionalLong getMaxProgress() {
        try {
            return OptionalLong.of(m_metadata.get().sheetMeta.sizeInBytes);
        } catch (final InterruptedException e) {
            Thread.currentThread().interrupt();
            return OptionalLong.empty();
        } catch (final ExecutionException e) {
            LOGGER.debug("Could not determine max progress", e);
            return OptionalLong.empty();
        }
    }

    @Override
    public long getProgress() {
        return m_currentProgress.get();
    }

//    public Map<String, Boolean> getSheetNames() {
//        try {
//            return m_metadata.get().sheetNames;
//        } catch (InterruptedException | ExecutionException e) {
//            return Collections.emptyMap();
//        }
//    }

    public WorkbookMetadata getWorkbookMetadata() throws InterruptedException, ExecutionException {
        return m_metadata.get();
    }

}
