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
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xls;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.OptionalLong;
import java.util.Set;
import java.util.function.Consumer;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;

/**
 * This class implements a read that uses the non-streaming API of Apache POI (usermodel API). It provides all
 * functionalities as it keeps the whole sheet in memory during reading and is able to read xls, xlsx, and xlsm file
 * formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSRead extends ExcelRead {

    private long m_numMaxRows;

    private int m_rowsRead;

    private Set<Integer> m_hiddenColumns;

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    public XLSRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
        final Consumer<Map<String, Boolean>> sheetNamesConsumer) throws IOException {
        super(path, config, sheetNamesConsumer);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @SuppressWarnings("resource") // workbook ownership handed to parser
    @Override
    public ExcelParserRunnable createParser(final File file) throws IOException {
        final var workbook = checkFileFormatAndCreateWorkbook(file);
        try {
            final var sheetNames = ExcelUtils.getSheetNames(workbook);
            if (m_sheetNamesConsumer != null) {
                m_sheetNamesConsumer.accept(sheetNames);
            }

            final var sheet = workbook.getSheet(getSelectedSheet(sheetNames));
            m_numMaxRows = sheet.getLastRowNum() + 1L;
            final FormulaEvaluator formulaEvaluator = m_config.getReaderSpecificConfig().isReevaluateFormulas()
                ? workbook.getCreationHelper().createFormulaEvaluator() : null;
            // ownership of workbook handed to parser
            return new XLSParserRunnable(this, m_config, workbook, sheet, use1904Windowing(workbook), formulaEvaluator);
        } catch (IOException e) {
            // if there was a problem prior to creating the parser, we have to close the workbook ourselves
            workbook.close();
            throw e;
        }
    }

    /**
     * Excel dates are encoded as a numeric offset of either 1900 or 1904.
     */
    private static boolean use1904Windowing(final Workbook workbook) {
        boolean date1904;
        if (workbook instanceof XSSFWorkbook xssfWorkbook) {
            date1904 = xssfWorkbook.isDate1904();
        } else if (workbook instanceof HSSFWorkbook hssfWorkbook) {
            date1904 = hssfWorkbook.getInternalWorkbook().isUsing1904DateWindowing();
        } else {
            // Probably unsupported
            date1904 = false;
        }
        return date1904;
    }

    /**
     * The same logic as used in {@link WorkbookFactory#create(InputStream, String)}. However, in order to prevent the
     * need to parse the message of the exception thrown there, we check ourselves and provide a more user-friendly
     * error message.
     */
    private Workbook checkFileFormatAndCreateWorkbook(final File file) throws IOException {
        final var type = FileMagic.valueOf(file);
        if (type != FileMagic.OLE2 && type != FileMagic.OOXML) {
            // will be caught and output with a user-friendly error message
            throw new NotOfficeXmlFileException("");
        }
        final var authModel = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
        final var password =
            CryptUtil.getPassword(authModel, m_config.getReaderSpecificConfig().getCredentialsProvider());
        try {
            return WorkbookFactory.create(file, password, true);
        } catch (final EncryptedDocumentException e) {
            if (password == null) {
                throw createPasswordProtectedFileException(e);
            }
            throw createPasswordIncorrectException(e);
        }
    }

    @Override
    public OptionalLong getMaxProgress() {
        return m_numMaxRows < 0 ? OptionalLong.empty() : OptionalLong.of(m_numMaxRows);
    }

    @Override
    public long getProgress() {
        return m_rowsRead;
    }

    @Override
    public Set<Integer> getHiddenColumns() {
        return m_hiddenColumns;
    }

    private class XLSParserRunnable extends ExcelParserRunnable {

        private final Workbook m_workbook;

        private final Sheet m_sheet;

        private final boolean m_use1904Windowing;

        private final FormulaEvaluator m_formulaEvaluator;

        private ExcelCell m_rowId;

        XLSParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,
            final Workbook workbook, final Sheet sheet,
            final boolean use1904Windowing, final FormulaEvaluator formulaEvaluator) {
            super(read, config);
            m_workbook = workbook;
            m_sheet = sheet;
            m_use1904Windowing = use1904Windowing;
            m_formulaEvaluator = formulaEvaluator;
            m_hiddenColumns = new HashSet<>();
        }

        @Override
        protected void parse() throws Exception {
            int lastNonEmptyRowIdx = -1;
            for (var i = 0; i <= m_sheet.getLastRowNum(); i++) {
                m_rowId = null;
                final var row = m_sheet.getRow(i);
                if (row != null) {
                    final boolean isHiddenRow = m_skipHiddenRows && row.getZeroHeight();
                    // parse the row
                    final List<ExcelCell> cells = parseRow(row);
                    // if all cells of the row are null, the row is empty
                    if (!isRowEmpty(cells) || m_rowId != null) {
                        outputEmptyRows(i - lastNonEmptyRowIdx - 1);
                        lastNonEmptyRowIdx = i;
                        // insert the row id at the beginning
                        insertRowIDAtBeginning(cells, m_rowId);
                        addToQueue(cells, isHiddenRow);
                    }
                }
            }
            outputEmptyRows(m_lastRowIdx - lastNonEmptyRowIdx);
        }

        private List<ExcelCell> parseRow(final Row row) {
            final List<ExcelCell> cells = new ArrayList<>();
            var numEmptyCells = 0;
            for (var j = 0; j < row.getLastCellNum(); j++) {
                final var cell = row.getCell(j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                final boolean isColRowID = isColRowID(j);
                if (isColRowID) {
                    m_rowId = cell == null ? null : parseCell(cell, cell.getCellType());
                }
                if (m_skipHiddenCols && m_sheet.isColumnHidden(j)) {
                    // we always need to add, not only for the first row, as a later row could have more columns
                    m_hiddenColumns.add(j);
                } else if (!isColRowID && isColIncluded(j)) {
                    if (cell == null) {
                        // by not adding null directly to the list but waiting for the next non-empty cell, we prevent
                        // columns with only "empty-but-formatted/styled" cells being appended (cells which had content
                        // before and were set to empty later might still be counted as cells by Excel and Apache POI)
                        numEmptyCells++;
                    } else {
                        addNullsToList(numEmptyCells, cells);
                        numEmptyCells = 0;
                        cells.add(parseCell(cell, cell.getCellType()));
                    }
                }
            }
            appendMissingCells(row.getLastCellNum(), cells);
            return cells;
        }

        private void addNullsToList(final int n, final List<?> list) {
            for (var i = 0; i < n; i++) {
                list.add(null);
            }
        }

        private void appendMissingCells(final int startColIdx, final List<?> list) {
            for (var j = startColIdx; j <= m_lastCol; j++) {
                if (m_skipHiddenCols && m_sheet.isColumnHidden(j)) {
                    m_hiddenColumns.add(j);
                } else if (!isColRowID(j) && isColIncluded(j)) {
                    list.add(null);
                }
            }
        }

        private ExcelCell parseCell(final Cell cell, final CellType cellType) {
            return switch (cellType) {
                case NUMERIC -> parseNumericOrDateCell(cell);
                case BOOLEAN -> new ExcelCell(KNIMECellType.BOOLEAN, cell.getBooleanCellValue());
                case STRING -> replaceStringWithMissing(cell.getStringCellValue()) ? null
                        : new ExcelCell(KNIMECellType.STRING, cell.getStringCellValue());
                case FORMULA -> parseFormulaCell(cell);
                case ERROR -> ExcelCellUtils.createErrorCell(m_config);
                // as we use MissingCellPolicy.RETURN_BLANK_AS_NULL, we should never get a BLANK
                default ->
                    throw new IllegalStateException("Unexpected cell type: " + cellType.toString());
            };
        }

        private ExcelCell parseNumericOrDateCell(final Cell cell) {
            final ExcelCell excelCell;
            if (DateUtil.isCellDateFormatted(cell)) {
                final var localDateTimeCellValue = cell.getLocalDateTimeCellValue();
                return ExcelCellUtils.createDateTimeExcelCell(localDateTimeCellValue, m_use1904Windowing);
            } else {
                final ExcelCell numericExcelCell = m_use15DigitsPrecision
                    ? ExcelCellUtils.createNumericExcelCellUsing15DigitsPrecision(cell.getNumericCellValue())
                    : ExcelCellUtils.createNumericExcelCell(cell.getNumericCellValue());
                if (cell.getCellType() == CellType.FORMULA) {
                    if (ExcelCellUtils.canBeBoolean(numericExcelCell)) {
                        excelCell = ExcelCellUtils.getIntOrBooleanCell(numericExcelCell, cell.toString());
                    } else {
                        excelCell = numericExcelCell;
                    }
                } else {
                    excelCell = numericExcelCell;
                }
            }
            return excelCell;
        }

        private ExcelCell parseEvaluatedNumericOrDateCellValue(final Cell cell, final CellValue cellValue) {
            final ExcelCell excelCell;
            if (DateUtil.isCellDateFormatted(cell)) {
                final var localDateTimeCellValue = DateUtil.getLocalDateTime(cellValue.getNumberValue());
                return ExcelCellUtils.createDateTimeExcelCell(localDateTimeCellValue, m_use1904Windowing);
            } else {
                excelCell = m_use15DigitsPrecision
                    ? ExcelCellUtils.createNumericExcelCellUsing15DigitsPrecision(cellValue.getNumberValue())
                    : ExcelCellUtils.createNumericExcelCell(cellValue.getNumberValue());
            }
            return excelCell;
        }

        private ExcelCell parseFormulaCell(final Cell cell) {
            final ExcelCell excelCell;
            if (m_formulaEvaluator != null) {
                excelCell = reevaluateAndParseFormulaCell(cell);
            } else {
                // get the type of the cached result and do a recursive call
                final var formulaResultCellType = cell.getCachedFormulaResultType();
                CheckUtils.checkState(formulaResultCellType != CellType.FORMULA,
                    "A formula cannot create another formula.");
                excelCell = parseCell(cell, formulaResultCellType);
            }
            return excelCell;
        }

        private ExcelCell reevaluateAndParseFormulaCell(final Cell cell) {
            final CellValue cellValue;
            try {
                cellValue = m_formulaEvaluator.evaluate(cell);
            } catch (Exception e) { // NOSONAR
                // as old node was catching for all Exceptions and I could not find documentation about the type of
                // exceptions thrown by FormulaEvaluator#evaluate, we do the same to avoid running into unexpected
                // exceptions (I encountered FormulaParseException and NotImplementedException)
                if (e.getClass() == RuntimeException.class
                    && e.getMessage().equals("Unexpected eval class (org.apache.poi.ss.formula.eval.BlankEval)")) {
                    return null;
                }
                return ExcelCellUtils.createErrorCell(m_config);
            }
            final var cellType = cellValue.getCellType();
            return switch (cellType) {
                case NUMERIC -> parseEvaluatedNumericOrDateCellValue(cell, cellValue);
                case BOOLEAN -> new ExcelCell(KNIMECellType.BOOLEAN, cellValue.getBooleanValue());
                case STRING -> replaceStringWithMissing(cellValue.getStringValue()) ? null
                        : new ExcelCell(KNIMECellType.STRING, cellValue.getStringValue());
                case ERROR -> ExcelCellUtils.createErrorCell(m_config);
                // as we use MissingCellPolicy.RETURN_BLANK_AS_NULL, we should never get a BLANK
                default ->
                    throw new IllegalStateException("Unexpected cell type: " + cellType.toString());
            };
        }

        @Override
        protected void closeResources() throws IOException {
            m_workbook.close();
        }

    }

}
