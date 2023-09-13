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
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.function.LongSupplier;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.AreaOfSheetToRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.CellConsumer;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils.ParsingInterruptedException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.VisibilityAwareRandomAccessible;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;

/**
 * A {@link SheetContentsHandler} for streamed parsing. If the handler is interrupted while outputting to the
 * {@link CellConsumer}, a runtime {@link ParsingInterruptedException} will be thrown.
 */
public class ExcelTableReaderSheetContentsHandler extends AbstractKNIMESheetContentsHandler {

    private int m_rowCount;

    private int m_currentRowIdx = -1;

    private int m_lastNonEmptyRowIdx = -1;

    private int m_currentCol = -1;

    private boolean m_currentRowIsHiddenAndSkipped;

    private final List<ExcelCell> m_row = new ArrayList<>();

    private ExcelCell m_rowId;

    private boolean m_endSheetCalled;


    private final KNIMEDataFormatter m_dataFormatter;

    /** The index of the fist included column. */
    protected final int m_firstCol;
    /** The index of the last included column or -1 if there is no max limit. */
    protected final int m_lastCol;
    /** The index of the last included row or -1 if there is no max limit. */
    protected final int m_lastRowIdx;
    /** The index of the row ID or -1 if no row ID column is set. */
    private final int m_rowIdIdx;
    /** True if hidden rows should be skipped. */
    private final boolean m_skipHiddenRows;
    /** True if hidden columns should be skipped. */
    protected final boolean m_skipHiddenCols;
    /** True if empty strings should be replaced by missings. */
    private final boolean m_replaceEmptyStringsWithMissings;

    private final boolean m_rawSettings;


    private final CellConsumer m_output;

    private final LongSupplier m_progress;

    private TableReadConfig<ExcelTableReaderConfig> m_config;

    /**
     * Constructor.
     *
     * @param dataFormatter the data formatter taking care of numeric and date&time cells
     */
    public ExcelTableReaderSheetContentsHandler(final CellConsumer output,
            final TableReadConfig<ExcelTableReaderConfig> config, final KNIMEDataFormatter dataFormatter,
            final LongSupplier progress) {
        m_output = output;
        m_config = config;
        m_dataFormatter = dataFormatter;
        m_progress = progress;

        final var excelConfig = config.getReaderSpecificConfig();
        m_firstCol = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
                ? ExcelUtils.getFirstColumnIdx(excelConfig.getReadFromCol()) : 0;
        m_lastCol = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
            ? ExcelUtils.getLastColumnIdx(excelConfig.getReadToCol()) : -1;
        m_lastRowIdx = excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL
                ? ExcelUtils.getLastRowIdx(excelConfig.getReadToRow()) : -1;
        m_rowIdIdx = config.useRowIDIdx() ? ExcelUtils.getRowIdColIdx(excelConfig.getRowIDCol()) : -1;
        m_skipHiddenCols = excelConfig.isSkipHiddenCols();
        m_skipHiddenRows = excelConfig.isSkipHiddenRows();
        m_replaceEmptyStringsWithMissings = excelConfig.isReplaceEmptyStringsWithMissings();
        m_rawSettings = excelConfig.isUseRawSettings();

    }

    @Override
    public void startRow(final int rowIdx) {
        m_currentRowIsHiddenAndSkipped = m_skipHiddenRows && isHiddenRow();
        m_currentRowIdx = rowIdx;
        m_currentCol = -1;
    }

    @Override
    public void endRow(final int rowIdx) {
        if (!ExcelUtils.isRowEmpty(m_row) || m_rowId != null) {
            // if there were empty rows in between two non-empty rows, output these (we need to make sure that
            // there is a non-empty row after an empty row)
            outputEmptyRows(m_currentRowIdx - m_lastNonEmptyRowIdx - 1);
            m_lastNonEmptyRowIdx = m_currentRowIdx;
            appendMissingCells();
            // insert the row id at the beginning
            insertRowIDAtBeginning(m_row, m_rowId);
            // add the non-empty row the the queue
            outputRow(m_row, m_currentRowIsHiddenAndSkipped);
        }
        m_rowId = null;
        m_row.clear();
    }


    /**
     * Inserts a row ID to the beginning of a row.
     *
     * @param row the row
     * @param rowId the row ID
     */
    private void insertRowIDAtBeginning(final List<ExcelCell> row, final ExcelCell rowId) {
        if (m_rowIdIdx >= 0) {
            row.add(0, rowId);
        }
    }

    /**
     * Adds the given number of empty {@link RandomAccessible}s to the blocking queue of the runnable.
     *
     * @param numMissingsRows the number of missing rows to add
     * @throws InterruptedException
     */
    private void outputEmptyRows(final int numMissingsRows) {
        for (var i = 0; i < numMissingsRows; i++) {
            final List<ExcelCell> cells = new ArrayList<>();
            if (m_rowIdIdx >= 0 && !m_rawSettings) {
                // make sure the empty row ID is added
                final ExcelCell cell = null;
                cells.add(cell);
            }
            outputRow(cells, false);
        }
    }

    private void outputRow(final List<ExcelCell> cells, final boolean isRowHidden) throws ParsingInterruptedException {
        if (m_rawSettings) {
            m_rowCount++;
            cells.add(0, new ExcelCell(KNIMECellType.INT, m_rowCount));
        }
        try {
            m_output.accept(VisibilityAwareRandomAccessible.createUnsafe(
                RandomAccessibleUtils.createFromArrayUnsafe(cells.toArray(new ExcelCell[0])), isRowHidden),
                m_progress.getAsLong());
        } catch (final InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new ParsingInterruptedException(e);
        }
    }

    @Override
    public void cell(String cellReference, final String formattedValue, final XSSFComment comment) {
        m_currentCol++;
        if (cellReference == null) {
            // gracefully handle missing CellRef here in a similar way as XSSFCell does.
            // according to the API description the cellReference argument shouldn't be null,
            // though it can be for files created by third party software (see e.g. AP-9380)
            cellReference = new CellAddress(m_currentRowIdx, m_currentCol).formatAsString();
        }

        final int currentColCount = m_currentCol;
        m_currentCol = new CellReference(cellReference).getCol();

        final Optional<ExcelCell> numericCell = m_dataFormatter.getAndResetExcelCell();
        final var excelCell = numericCell.orElseGet(() -> createExcelCellFromStringValue(formattedValue));
        handleMissingCells(cellReference, currentColCount);
        if (!isColHiddenAndSkipped(m_currentCol, getHiddenCols())) {
            if (isColRowID(m_currentCol)) {
                // check if cells were missing and insert nulls if so
                m_rowId = excelCell;
            } else if (isColIncluded(m_currentCol)) {
                // check if cells were missing and insert nulls if so
                m_row.add(excelCell);
            }
        } else if (isColRowID(m_currentCol)) {
            m_rowId = excelCell;
        }
    }

    @Override
    public void endSheet() {
        if (!m_endSheetCalled) {
            outputEmptyRows(m_lastRowIdx - m_lastNonEmptyRowIdx);
            m_endSheetCalled = true;
        }
    }

    private ExcelCell createExcelCellFromStringValue(final String formattedValue) {
        final ExcelCell excelCell;
        final KNIMEXSSFDataType cellType = getCellType();
        if (cellType == null) {
            throw new IllegalStateException("Coding error: encountered unexpected cell type.");
        }
        switch (cellType) {
            case BOOLEAN:
                excelCell = new ExcelCell(KNIMECellType.BOOLEAN, formattedValue.equals("TRUE"));
                break;
            case ERROR:
                excelCell = ExcelCellUtils.createErrorCell(m_config);
                break;
            case FORMULA:
                excelCell = replaceStringWithMissing(formattedValue) ? null
                    : new ExcelCell(KNIMECellType.STRING, formattedValue);
                break;
            case STRING:
                excelCell = replaceStringWithMissing(formattedValue) ? null
                    : new ExcelCell(KNIMECellType.STRING, formattedValue);
                break;
            case NUMBER_OR_DATE:
                // special case, see AP-16880; required for cells that are processed by
                // org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler line 372
                // (cells with data type NUMBER but empty string as value)
                excelCell = null;
                break;
            default:
                throw new IllegalStateException("Unexpected data type: " + cellType);
        }
        return excelCell;
    }

    @Override
    public void headerFooter(final String text, final boolean isHeader, final String tagName) {
        // do nothing, we do not treat header and footer
    }

    private void handleMissingCells(final String cellReference, final int currentColCount) {
        final int thisCol = new CellReference(cellReference).getCol();
        final int missedCols = thisCol - currentColCount;
        for (int currentCol = currentColCount; currentCol < currentColCount + missedCols; currentCol++) {
            if (isColIncluded(currentCol) && !isColHiddenAndSkipped(currentCol, getHiddenCols())
                && !isColRowID(currentCol)) {
                m_row.add(null);
            }
        }
    }

    private void appendMissingCells() {
        m_currentCol++;
        for (int currentCol = m_currentCol; currentCol <= m_lastCol; currentCol++) {
            if (isColIncluded(currentCol) && !isColHiddenAndSkipped(currentCol, getHiddenCols())
                && !isColRowID(currentCol)) {
                m_row.add(null);
            }
        }
    }

    private boolean isColHiddenAndSkipped(final int col, final Set<Integer> hiddenCols) {
        return m_skipHiddenCols && hiddenCols.contains(col);
    }


    /**
     * Returns true if the column is the row id.
     *
     * @param col the column index
     * @return whether the column is the row id
     */
    private boolean isColRowID(final int col) {
        return col == m_rowIdIdx;
    }

    /**
     * Returns true if the column is included.
     *
     * @param col the column index
     * @return whether the column is included
     */
    private boolean isColIncluded(final int col) {
        return col >= m_firstCol && (m_lastCol < 0 || col <= m_lastCol);
    }


    /**
     * Returns whether the passed string should be replaced with a missing value which is {@code true} if
     * {@link String#trim()} returns an empty string and the user has configured the node to replace such strings with
     * missing values.
     *
     * @param string the string to test
     * @return whether the passed string should be replaced with a missing value
     */
    private boolean replaceStringWithMissing(final String string) {
        return m_replaceEmptyStringsWithMissings && string.trim().isEmpty();
    }
}