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
 *   Dec 3, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;

/**
 * Abstract class extending {@link ExcelParserRunnable} that is used for parsing files in streamed fashion.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class AbstractStreamedParserRunnable extends ExcelParserRunnable {

    private final TableReadConfig<ExcelTableReaderConfig> m_config;

    private Set<Integer> m_hiddenColumns = new HashSet<>();

    /**
     * Constructor.
     *
     * @param read the excel read
     * @param config the config
     */
    protected AbstractStreamedParserRunnable(final ExcelRead read,
        final TableReadConfig<ExcelTableReaderConfig> config) {
        super(read, config);
        m_config = config;
    }

    /**
     * Returns the indexes of the hidden columns.
     *
     * @return the indexes of the hidden columns in a {@link Set}
     */
    Set<Integer> getHiddenColumns() {
        return m_hiddenColumns;
    }

    /**
     * A {@link SheetContentsHandler} for streamed parsing.
     */
    protected class ExcelTableReaderSheetContentsHandler extends AbstractKNIMESheetContentsHandler {

        private final KNIMEDataFormatter m_dataFormatter;

        private int m_currentRowIdx = -1;

        private int m_lastNonEmptyRowIdx = -1;

        private int m_currentCol = -1;

        private boolean m_currentRowIsHiddenAndSkipped;

        private final List<ExcelCell> m_row = new ArrayList<>();

        private ExcelCell m_rowId;

        private boolean m_endSheetCalled;

        /**
         * Constructor.
         *
         * @param dataFormatter the data formatter taking care of numeric and date&time cells
         */
        public ExcelTableReaderSheetContentsHandler(final KNIMEDataFormatter dataFormatter) {
            m_dataFormatter = dataFormatter;
        }

        @Override
        public void startRow(final int rowIdx) {
            m_currentRowIsHiddenAndSkipped = m_skipHiddenRows && isHiddenRow();
            m_currentRowIdx = rowIdx;
            m_currentCol = -1;
        }

        @Override
        public void endRow(final int rowIdx) {
            if (!isRowEmpty(m_row) || m_rowId != null) {
                // if there were empty rows in between two non-empty rows, output these (we need to make sure that
                // there is a non-empty row after an empty row)
                outputEmptyRows(m_currentRowIdx - m_lastNonEmptyRowIdx - 1);
                m_lastNonEmptyRowIdx = m_currentRowIdx;
                m_hiddenColumns = getHiddenCols();
                appendMissingCells();
                // insert the row id at the beginning
                insertRowIDAtBeginning(m_row, m_rowId);
                // add the non-empty row the the queue
                addToQueue(m_row, m_currentRowIsHiddenAndSkipped);
            }
            m_rowId = null;
            m_row.clear();
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
    }
}