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
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.core.node.NodeLogger;
import org.knime.core.util.ThreadUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;
import org.knime.filehandling.core.node.table.reader.read.Read;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * This class implements a read that uses the streaming API of Apache POI (eventmodel API) and can read xlsx and xlsm
 * file formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSXRead implements Read<String> {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSXRead.class);

    private static final int BLOCKING_QUEUE_SIZE = 100;

    private static final AtomicLong CACHED_THREAD_POOL_INDEX = new AtomicLong();

    /** Threadpool that can be used to parse new sheets. */
    private static final ExecutorService CACHED_THREAD_POOL = Executors.newCachedThreadPool(
        r -> ThreadUtils.threadWithContext(r, "KNIME-XLSX-Parser-" + CACHED_THREAD_POOL_INDEX.getAndIncrement()));

    /** The poison pill put into the blocking to indicate end of parsing. */
    private static final RandomAccessible<String> POISON_PILL = RandomAccessibleUtils.createFromArray("POISON");

    /** The path of the underlying source. */
    private final Path m_path;

    /** Queue that uses the TRF to consume rows produced by the parser thread. */
    private final BlockingQueue<RandomAccessible<String>> m_queueRandomAccessibles =
        new ArrayBlockingQueue<>(BLOCKING_QUEUE_SIZE);

    /** The thread running the parser. */
    private final Future<?> m_parserThread;

    /** Atomic reference used to communicate exceptions between parser thread and main thread. */
    private final AtomicReference<Throwable> m_throwableDuringParsing = new AtomicReference<>();

    /** Iterator collecting RandomAccessibles from the parser thread. */
    private final RandomAccessibleIterator m_randomAccessibleIterator;

    private final long m_sheetSize;

    private final OPCPackage m_opc;

    private final InputStream m_inputStream;

    private final CountingInputStream m_countingSheetStream;

    /**
     * Constructor
     *
     * @param path the path of the file to read
     * @param config the Excel table reader configuration.
     * @throws IOException if a stream can not be created from the provided file.
     */
    @SuppressWarnings("resource") // we are going to close all the resources in #close
    public XLSXRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        m_path = path;
        m_randomAccessibleIterator = new RandomAccessibleIterator();

        m_inputStream = FSFiles.newInputStream(path);
        try {
            m_opc = OPCPackage.open(m_inputStream);

            // take the first sheet
            final XSSFReader xssfReader = new XSSFReader(m_opc);
            final SheetIterator sheetsData = (SheetIterator)xssfReader.getSheetsData();
            if (!sheetsData.hasNext()) {
                throw new IOException("Selected file does not have any sheet.");
            }
            m_countingSheetStream = new CountingInputStream(sheetsData.next());
            m_sheetSize = m_countingSheetStream.available();

            // create the parser
            final XMLReader sheetParser = XMLHelper.newXMLReader();
            final ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(m_opc, false);
            final ExcelTableReaderSheetContentsHandler sheetContentsHandler =
                new ExcelTableReaderSheetContentsHandler();
            final DataFormatter dataFormatter = new DataFormatter();

            sheetParser.setContentHandler(new XSSFSheetXMLHandler(xssfReader.getStylesTable(), null, sharedStringsTable,
                sheetContentsHandler, dataFormatter, false));

            // create the runnable that will execute the parsing
            final Runnable runnable = () -> {
                try {
                    sheetParser.parse(new InputSource(m_countingSheetStream));
                    m_queueRandomAccessibles.put(POISON_PILL);
                    LOGGER.debug("Thread parsing an xlsx sheet finished successfully.");
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    closeAndSwallowExceptions();
                } catch (Throwable e) { // NOSONAR we want to catch any Throwable
                    m_throwableDuringParsing.set(e);
                    // cannot do anything with the IO exceptions besides logging
                    closeAndSwallowExceptions();
                }
            };

            // create and start the thread
            m_parserThread = CACHED_THREAD_POOL.submit(ThreadUtils.runnableWithContext(runnable));
            LOGGER.debug("Thread parsing an xlsx sheet started.");
        } catch (InvalidOperationException | OpenXML4JException | SAXException | ParserConfigurationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private void closeAndSwallowExceptions() {
        try {
            close();
        } catch (IOException ioe) {
            LOGGER.debug(ioe);
        }
    }

    @Override
    public RandomAccessible<String> next() throws IOException {
        final boolean hasNext = m_randomAccessibleIterator.hasNext();
        if (m_throwableDuringParsing.get() != null) {
            close();
        }
        return hasNext ? m_randomAccessibleIterator.next() : null;
    }

    @Override
    public OptionalLong getEstimatedSizeInBytes() {
        return m_sheetSize < 0 ? OptionalLong.empty() : OptionalLong.of(m_sheetSize);
    }

    @Override
    public long readBytes() {
        return m_countingSheetStream.getByteCount();
    }

    @Override
    public Optional<Path> getPath() {
        return Optional.ofNullable(m_path);
    }

    @Override
    public void close() throws IOException {
        // cancel the thread
        m_parserThread.cancel(true);
        // free the queue so that any put call by the canceled thread is not blocking the cancellation
        m_queueRandomAccessibles.clear();
        // as we just cleared the queue, we can use #add to insert the poison pill
        m_queueRandomAccessibles.add(POISON_PILL);

        final Throwable throwable = m_throwableDuringParsing.get();
        // close resources
        try {
            m_countingSheetStream.close();
            m_inputStream.close();
            m_opc.close();
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
     * Iterator that collects all parsed rows. If an exception occurred during parsing or the parsing is finished,
     * {@link #hasNext()} will return {@code false}. Otherwise, it will wait for more rows becoming available to iterate
     * over.
     */
    private class RandomAccessibleIterator implements Iterator<RandomAccessible<String>> {

        private final LinkedList<RandomAccessible<String>> m_randomAccessibles = new LinkedList<>();

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
        public RandomAccessible<String> next() {
            if (!hasNext()) {
                throw new NoSuchElementException("No more RandomAccessible available.");
            }
            return m_randomAccessibles.poll();
        }

    }

    private class ExcelTableReaderSheetContentsHandler implements SheetContentsHandler {

        private int m_currentRow = -1;

        private int m_currentCol = -1;

        private final ArrayList<String> m_row = new ArrayList<>();

        /*
         *  TODO this works only if short rows are allowed which is right now always enabled (old behavior as well).
         *  if we want to offer short rows as a setting, we need to revise the below method.
         */
        private void outputEmptyRows(final int numMissingsRows) {
            for (int i = 0; i < numMissingsRows; i++) {
                m_currentRow++;
                endRow(m_currentRow);
            }
        }

        @Override
        public void startRow(final int rowNum) {
            // If there were empty rows, output these
            outputEmptyRows(rowNum - m_currentRow - 1);

            m_currentRow = rowNum;
            m_currentCol = -1;
        }

        @Override
        public void endRow(final int rowNum) {
            try {
                m_queueRandomAccessibles.put(RandomAccessibleUtils.createFromArrayUnsafe(m_row.toArray(new String[0])));
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                LOGGER.debug("Thread parsing an xlsx sheet interrupted.");
                // throw a runtime exception as the interrupting the thread does not stop it
                throw new ParsingInterruptedException();
            }
            m_row.clear();
        }

        @Override
        public void cell(String cellReference, final String formattedValue, final XSSFComment comment) {
            m_currentCol++;
            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if (cellReference == null) {
                cellReference = new CellAddress(m_currentRow, m_currentCol).formatAsString();
            }

            // check if cells were missing and insert null if so
            int thisCol = new CellReference(cellReference).getCol();
            int missedCols = thisCol - m_currentCol;
            for (int i = 0; i < missedCols; i++) {
                m_row.add(null);
            }
            m_currentCol += missedCols;

            // add the cell value to the row
            m_row.add(formattedValue);
        }

        @Override
        public void headerFooter(final String text, final boolean isHeader, final String tagName) {
            // TODO do nothing? (that's at least old behavior) -> we will see later
        }

    }

    /**
     * Exception to be thrown when the thread that parses the sheet should be interrupted.
     *
     * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
     */
    private static class ParsingInterruptedException extends RuntimeException {

        private static final long serialVersionUID = 1L;

    }
}
