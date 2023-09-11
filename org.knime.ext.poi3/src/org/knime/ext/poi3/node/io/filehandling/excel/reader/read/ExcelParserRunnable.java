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
 *   Oct 22, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.AreaOfSheetToRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils.ParsingInterruptedException;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;
import org.xml.sax.SAXException;

/**
 * Abstract implementation of a {@link Runnable} that parses Excel files and adds the parsed rows as
 * {@link RandomAccessible} to the queue of the provided {@link ExcelRead}.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class ExcelParserRunnable implements Runnable {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelParserRunnable.class);

    private final ExcelRead m_read;

    /** True if 15 digits precision is used (that's what Excel is using). */
    protected final boolean m_use15DigitsPrecision;

    /** True if hidden columns should be skipped. */
    protected final boolean m_skipHiddenCols;

    /** True if hidden rows should be skipped. */
    protected final boolean m_skipHiddenRows;

    /** True if empty strings should be replaced by missings. */
    private final boolean m_replaceEmptyStringsWithMissings;

    /** The index of the row ID or -1 if no row ID column is set. */
    protected final int m_rowIdIdx;

    /** The index of the fist included column. */
    protected final int m_firstCol;

    /** The index of the last included column or -1 if there is no max limit. */
    protected final int m_lastCol;

    /** The index of the last included row or -1 if there is no max limit. */
    protected final int m_lastRowIdx;

    private final boolean m_rawSettings;

    private int m_rowCount;

    /**
     * Constructor.
     *
     * @param read the {@link ExcelRead}
     * @param config the config
     */
    protected ExcelParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config) {
        m_read = read;
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        m_use15DigitsPrecision = excelConfig.isUse15DigitsPrecision();
        m_skipHiddenCols = excelConfig.isSkipHiddenCols();
        m_skipHiddenRows = excelConfig.isSkipHiddenRows();
        m_replaceEmptyStringsWithMissings = excelConfig.isReplaceEmptyStringsWithMissings();
        m_rowIdIdx = config.useRowIDIdx() ? ExcelUtils.getRowIdColIdx(excelConfig.getRowIDCol()) : -1;
        m_firstCol = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
            ? ExcelUtils.getFirstColumnIdx(excelConfig.getReadFromCol()) : 0;
        m_lastCol = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
            ? ExcelUtils.getLastColumnIdx(excelConfig.getReadToCol()) : -1;
        m_lastRowIdx = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
                ? ExcelUtils.getLastRowIdx(excelConfig.getReadToRow()) : -1;
        m_rawSettings = excelConfig.isUseRawSettings();
    }

    @Override
    public void run() {
        LOGGER.debug("Thread parsing an Excel spreadsheet started.");
        try {
            parse();
            indicateFinished();
            LOGGER.debug("Thread parsing an Excel spreadsheet finished successfully.");
        } catch (final ParsingInterruptedException e) { // NOSONAR ignore, they are thrown by us to indicate parsing should stop
            // thrown e.g. in ExcelParserRunnable#addToQueue, ExcelUtils.IsEmpty#cell
            closeAndSwallowExceptions();
//        } catch (final ClosedByInterruptException e) { // NOSONAR different line for documentation reasons
//            // parser was interrupted while blocked in I/O operation and channel was closed
//            closeAndSwallowExceptions();
        } catch (final IOException | SAXException | InvalidFormatException e) {
            // rethrow later by Read
            m_read.setThrowable(e);
            closeAndSwallowExceptions();
        }
    }

    private void indicateFinished() {
        try {
            m_read.addToQueue(ExcelRead.POISON_PILL);
        } catch (final InterruptedException e) {
            // while adding to queue
            Thread.currentThread().interrupt();
            closeAndSwallowExceptions();
        }
    }

    private void closeAndSwallowExceptions() {
        try {
            m_read.close();
        } catch (IOException ioe) {
            LOGGER.debug("Exception while closing read due to parser problem.", ioe);
        }
    }

    /**
     * Start parsing.
     *
     * @throws IOException when an I/O problem occurs
     * @throws InvalidFormatException the file does not adhere to the Excel format to be parsed
     * @throws SAXException general SAX parsing exception for XML-based Excel files
     *
     */
    protected abstract void parse() throws IOException, InvalidFormatException, SAXException;

    /**
     * Adds the {@link RandomAccessible} to the blocking queue of the runnable.
     *
     * @param cells a list of {@link ExcelCell}s
     * @param isRowHidden if the row is hidden (and hidden rows should be skipped)
     */
    protected void addToQueue(final List<ExcelCell> cells, final boolean isRowHidden) {
        try {
            if (m_rawSettings) {
                m_rowCount++;
                cells.add(0, new ExcelCell(KNIMECellType.INT, m_rowCount));
            }
            m_read.addToQueue(VisibilityAwareRandomAccessible.createUnsafe(
                RandomAccessibleUtils.createFromArrayUnsafe(cells.toArray(new ExcelCell[0])), isRowHidden));
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            LOGGER.debug("Thread parsing an Excel spreadsheet interrupted.");
            // throw a runtime exception as the interrupting the thread does not stop it
            throw new ParsingInterruptedException();
        }
    }

    /**
     * Adds the given number of empty {@link RandomAccessible}s to the blocking queue of the runnable.
     *
     * @param numMissingsRows the number of missing rows to add
     */
    protected void outputEmptyRows(final int numMissingsRows) {
        for (var i = 0; i < numMissingsRows; i++) {
            final List<ExcelCell> cells = new ArrayList<>();
            if (m_rowIdIdx >= 0 && !m_rawSettings) {
                // make sure the empty row ID is added
                final ExcelCell cell = null;
                cells.add(cell);
            }
            addToQueue(cells, false);
        }
    }

    /**
     * Inserts a row ID to the beginning of a row.
     *
     * @param row the row
     * @param rowId the row ID
     */
    protected void insertRowIDAtBeginning(final List<ExcelCell> row, final ExcelCell rowId) {
        if (m_rowIdIdx >= 0) {
            row.add(0, rowId);
        }
    }

    /**
     * Returns true if the row is empty, i.e., it contains only nulls or no elements.
     *
     * @param row the row
     * @return whether the row is empty
     */
    protected boolean isRowEmpty(final List<ExcelCell> row) {
        return row.isEmpty() || row.stream().allMatch(Objects::isNull);
    }

    /**
     * Returns true if the column is included.
     *
     * @param col the column index
     * @return whether the column is included
     */
    protected boolean isColIncluded(final int col) {
        return col >= m_firstCol && (m_lastCol < 0 || col <= m_lastCol);
    }

    /**
     * Returns true if the column is the row id.
     *
     * @param col the column index
     * @return whether the column is the row id
     */
    protected boolean isColRowID(final int col) {
        return col == m_rowIdIdx;
    }

    /**
     * Returns whether the passed string should be replaced with a missing value which is {@code true} if
     * {@link String#trim()} returns an empty string and the user has configured the node to replace such strings with
     * missing values.
     *
     * @param string the string to test
     * @return whether the passed string should be replaced with a missing value
     */
    protected boolean replaceStringWithMissing(final String string) {
        return m_replaceEmptyStringsWithMissings && string.trim().isEmpty();
    }

}
