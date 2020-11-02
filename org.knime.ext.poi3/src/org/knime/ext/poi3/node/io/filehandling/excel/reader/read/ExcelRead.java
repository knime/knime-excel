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

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.knime.core.util.ThreadUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.connections.FSFiles;
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

    /** The String to be inserted when an error cell is encountered. */
    protected static final String XL_EVAL_ERROR = "#XL_EVAL_ERROR#";

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

    /** The input stream of the file. */
    private final InputStream m_inputStream;

    /**
     * Constructor
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if a stream can not be created from the provided file.
     */
    protected ExcelRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        m_path = path;
        m_config = config;
        m_inputStream = FSFiles.newInputStream(path);
        startParserThread();
    }

    /**
     * Creates a {@code ExcelParserRunnable}.
     *
     * @param inputStream the file input stream
     * @return the {@code ExcelParserRunnable}
     * @throws IOException if an I/O exception occurs
     */
    protected abstract ExcelParserRunnable createParser(final InputStream inputStream) throws IOException;

    /**
     * Closes resources.
     *
     * @throws IOException if an I/O exception occurs
     */
    protected abstract void closeResources() throws IOException;

    @Override
    public final RandomAccessible<ExcelCell> next() throws IOException {
        // check and return next element
        final boolean hasNext = m_randomAccessibleIterator.hasNext();
        if (m_throwableDuringParsing.get() != null) {
            close();
        }
        return hasNext ? m_randomAccessibleIterator.next() : null;
    }

    private void startParserThread() throws IOException {
        try {
            // create and start the thread
            final ExcelParserRunnable runnable = createParser(m_inputStream);
            m_parserThread = CACHED_THREAD_POOL.submit(ThreadUtils.runnableWithContext(runnable));
        } catch (InvalidOperationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    @Override
    public final Optional<Path> getPath() {
        return Optional.of(m_path);
    }

    @Override
    public final void close() throws IOException {
        // cancel the thread
        if (m_parserThread != null) {
            m_parserThread.cancel(true);
        }
        // free the queue so that any put call by the canceled thread is not blocking the cancellation
        m_queueRandomAccessibles.clear();
        // as we just cleared the queue, we can use #add to insert the poison pill
        m_queueRandomAccessibles.add(POISON_PILL);

        final Throwable throwable = m_throwableDuringParsing.get();
        // close resources
        try {
            closeResources();
            if (m_inputStream != null) {
                m_inputStream.close();
            }
        } catch (IOException e) {
            // an exception during parsing has priority
            if (throwable == null) {
                throw e;
            }
        }

        // if an exception occurred during parsing, throw it
        if (throwable != null) {
            if (throwable instanceof RuntimeException) {
                throw (RuntimeException)throwable;
            } else if (throwable instanceof Error) {
                throw (Error)throwable;
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

        private boolean m_encounteredPoisonPill = false;

        @Override
        public boolean hasNext() {
            if (m_encounteredPoisonPill || m_throwableDuringParsing.get() != null) {
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

}
