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

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.OptionalLong;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;

/**
 * This class implements a read that uses the non-streaming API of Apache POI (usermodel API). It provides all
 * functionalities as it keeps the whole sheet in memory during reading and is able to read xls, xlsx, and xlsm file
 * formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSRead extends ExcelRead {

    private long m_numMaxRows;

    private int m_rowsRead = 0;

    private Workbook m_workbook;

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    public XLSRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        super(path, config);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @Override
    public ExcelParserRunnable createParser(final InputStream inputStream) throws IOException {
        try {
            final BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);
            m_workbook = checkFileFormatAndCreateWorkbook(bufferedInputStream);
            if (m_workbook.getNumberOfSheets() < 1) {
                throw new IOException("Selected file does not contain any sheet.");
            }
            // take first sheet
            final Sheet sheet = m_workbook.getSheetAt(0);
            m_numMaxRows = sheet.getLastRowNum() + 1L;

            return new XLSParserRunnable(this, m_config, sheet, use1904Windowing(m_workbook));
        } catch (InvalidOperationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    /**
     * Excel dates are encoded as a numeric offset of either 1900 or 1904.
     */
    private static boolean use1904Windowing(final Workbook workbook) {
        boolean date1904;
        if (workbook instanceof XSSFWorkbook) {
            final XSSFWorkbook xssfWorkbook = (XSSFWorkbook)workbook;
            date1904 = xssfWorkbook.isDate1904();
        } else if (workbook instanceof HSSFWorkbook) {
            final HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
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
    private static Workbook checkFileFormatAndCreateWorkbook(final BufferedInputStream inputStream) throws IOException {
        switch (FileMagic.valueOf(inputStream)) {
            case OLE2:
            case OOXML:
                break;
            default:
                // will be caught and output with a user-friendly error message
                throw new NotOfficeXmlFileException("");
        }
        return WorkbookFactory.create(inputStream);
    }

    @Override
    public OptionalLong getEstimatedSizeInBytes() {
        return m_numMaxRows < 0 ? OptionalLong.empty() : OptionalLong.of(m_numMaxRows);
    }

    @Override
    public long readBytes() {
        return m_rowsRead;
    }

    @Override
    public void closeResources() throws IOException {
        if (m_workbook != null) {
            m_workbook.close();
        }
    }

    private static class XLSParserRunnable extends ExcelParserRunnable {

        private Sheet m_sheet;

        private final boolean m_use1904Windowing;

        XLSParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config, final Sheet sheet,
            final boolean use1904Windowing) {
            super(read, config);
            m_sheet = sheet;
            m_use1904Windowing = use1904Windowing;
        }

        @Override
        protected void parse() throws Exception {
            int numEmptyRows = 0;
            for (int i = 0; i <= m_sheet.getLastRowNum(); i++) {
                final Row row = m_sheet.getRow(i);
                if (row == null) {
                    // empty row
                    numEmptyRows++;
                    continue;
                }
                // parse the row
                final List<ExcelCell> cells = new ArrayList<>();
                int numEmptyCells = 0;
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    final Cell cell = row.getCell(j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
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
                // check if all cells of the row are which means the row is empty
                if (cells.stream().noneMatch(Objects::nonNull)) {
                    // by not adding empty rows directly to the queue but waiting for the next non-empty row, we prevent
                    // rows with only "empty-but-formatted/styled" cells being added (cells which had content
                    // before and were set to empty later might still be counted as cells by Excel and Apache POI)
                    numEmptyRows++;
                } else {
                    outputEmptyRows(numEmptyRows);
                    numEmptyRows = 0;
                    addToQueue(RandomAccessibleUtils.createFromArrayUnsafe(cells.toArray(new ExcelCell[0])));
                }
            }
        }

        private static void addNullsToList(final int n, final List<?> list) {
            for (int i = 0; i < n; i++) {
                list.add(null);
            }
        }

        private ExcelCell parseCell(final Cell cell, final CellType cellType) {
            final ExcelCell excelCell;
            switch (cellType) {
                case NUMERIC:
                    excelCell = parseNumericOrDateCell(cell);
                    break;
                case BOOLEAN:
                    excelCell = new ExcelCell(KNIMECellType.BOOLEAN, cell.getBooleanCellValue());
                    break;
                case STRING:
                    excelCell = new ExcelCell(KNIMECellType.STRING, cell.getStringCellValue());
                    break;
                case FORMULA:
                    excelCell = parseFormulaCell(cell);
                    break;
                case ERROR:
                    // TODO default behavior of the old reader is to return "#XL_EVAL_ERROR#", treat later in AP-15398
                    excelCell = new ExcelCell(KNIMECellType.STRING, XL_EVAL_ERROR);
                    break;
                case BLANK:
                    // as we use MissingCellPolicy.RETURN_BLANK_AS_NULL, we should never get a BLANK
                default:
                    throw new IllegalStateException("Unexpected cell type: " + cellType.toString());
            }
            return excelCell;
        }

        private ExcelCell parseNumericOrDateCell(final Cell cell) {
            final ExcelCell excelCell;
            if (DateUtil.isCellDateFormatted(cell)) {
                final LocalDateTime localDateTimeCellValue = cell.getLocalDateTimeCellValue();
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

        private ExcelCell parseFormulaCell(final Cell cell) {
            // get the type of the cached result and do a recursive call
            final CellType formulaResultCellType = cell.getCachedFormulaResultType();
            CheckUtils.checkState(formulaResultCellType != CellType.FORMULA,
                "A formula cannot create another formula.");
            return parseCell(cell, formulaResultCellType);
        }

    }

}
