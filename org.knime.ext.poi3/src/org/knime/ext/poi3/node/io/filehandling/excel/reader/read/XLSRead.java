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

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.OptionalLong;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
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
    }

    @Override
    public ExcelParserRunnable createParser(final InputStream inputStream) throws IOException {
        try {
            final BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);
            checkFileFormat(bufferedInputStream);
            m_workbook = WorkbookFactory.create(bufferedInputStream);
            if (m_workbook.getNumberOfSheets() < 1) {
                throw new IOException("Selected file does not have any sheet.");
            }
            // take first sheet
            final Sheet sheet = m_workbook.getSheetAt(0);
            m_numMaxRows = sheet.getLastRowNum() + 1L;

            return new XLSParserRunnable(this, sheet, getTimeOnlyYearIdentifier(m_workbook));
        } catch (InvalidOperationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    /**
     * Cells with only time but no date value have date set to 1899-12-31 or 1904-01-01 depending on if the date is
     * using 1904 windowing.
     */
    private static int getTimeOnlyYearIdentifier(final Workbook workbook) {
        boolean date1904;
        if (workbook instanceof XSSFWorkbook) {
            final XSSFWorkbook xssfWorkbook = (XSSFWorkbook)workbook;
            date1904 = xssfWorkbook.isDate1904();
        } else if (workbook instanceof HSSFWorkbook) {
            final HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
            date1904 = hssfWorkbook.getInternalWorkbook().isUsing1904DateWindowing();
        } else {
            //Probably unsupported
            date1904 = false;
        }
        return date1904 ? 1904 : 1899;
    }

    /**
     * The same logic as used in {@link WorkbookFactory#create(InputStream, String)}. However, in order to prevent the
     * need to parse the message of the exception thrown there, we check ourselves and provide a more user-friendly
     * error message.
     */
    private void checkFileFormat(final BufferedInputStream inputStream) throws IOException {
        switch (FileMagic.valueOf(inputStream)) {
            case OLE2:
                break;
            case OOXML:
                break;
            default:
                throw unsupportedFileFormatException(null);
        }
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
        m_workbook.close();
    }

    private static class XLSParserRunnable extends ExcelParserRunnable {

        private Sheet m_sheet;

        private final DataFormatter m_dataFormatter = new DataFormatter();

        private final int m_timeOnlyYearIdentifier;

        XLSParserRunnable(final ExcelRead read, final Sheet sheet, final int timeOnlyYearIdentifier) {
            super(read);
            m_sheet = sheet;
            m_timeOnlyYearIdentifier = timeOnlyYearIdentifier;
        }

        @Override
        protected void parse() throws Exception {
            for (int i = 0; i <= m_sheet.getLastRowNum(); i++) {
                final Row row = m_sheet.getRow(i);
                if (row == null) {
                    // empty row
                    addToQueue(RandomAccessibleUtils.createFromArrayUnsafe());
                    continue;
                }
                final List<String> cells = new ArrayList<>();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    final Cell cell = row.getCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cells.add(parseCell(cell, cell.getCellType()));
                }
                addToQueue(RandomAccessibleUtils.createFromArrayUnsafe(cells.toArray(new String[0])));
            }
        }

        private String parseCell(final Cell cell, final CellType cellType) {
            final String stringValue;
            switch (cellType) {
                case NUMERIC:
                    stringValue = parseNumericCell(cell);
                    break;
                case BLANK:
                    stringValue = null;
                    break;
                case BOOLEAN:
                    final boolean booleanCellValue = cell.getBooleanCellValue();
                    stringValue = String.valueOf(booleanCellValue);
                    break;
                case STRING:
                    stringValue = cell.getStringCellValue();
                    break;
                case FORMULA:
                    final CellType formulaResultCellType = cell.getCachedFormulaResultType();
                    CheckUtils.checkState(formulaResultCellType != CellType.FORMULA,
                        "A formula cannot create another formula.");
                    return parseCell(cell, formulaResultCellType);
                case ERROR:
                    // TODO default behavior of the old reader is to return "#XL_EVAL_ERROR#", treat later in AP-15398
                    if (cell.getCellType() == CellType.FORMULA) {
                        stringValue = "#XL_EVAL_ERROR#";
                    } else {
                        stringValue = "ERROR:" + m_dataFormatter.formatCellValue(cell);
                    }
                    break;
                default:
                    throw new IllegalStateException("Unexpected cell type: " + cellType.toString());
            }
            return stringValue;
        }

        private String parseNumericCell(final Cell cell) {
            final String stringValue;
            if (DateUtil.isCellDateFormatted(cell)) {
                final LocalDateTime localDateTimeCellValue = cell.getLocalDateTimeCellValue();
                if (localDateTimeCellValue.getYear() == m_timeOnlyYearIdentifier) {
                    stringValue = localDateTimeCellValue.toLocalTime().toString();
                } else {
                    // We cannot know if only a date or date&time is set. Only-date values have time set
                    // to 00:00 which could be midnight for date&time.
                    // TODO if all values in a column have 00:00, we can assume date-only -> part of type hierarchy?
                    stringValue = localDateTimeCellValue.toString();
                }
            } else {
                if (cell.getCellType() == CellType.FORMULA) {
                    // format as raw cell in order to apply the formatter to cached formula result and not on the
                    // formula itself
                    final CellStyle cellStyle = cell.getCellStyle();
                    stringValue = m_dataFormatter.formatRawCellContents(cell.getNumericCellValue(),
                        cellStyle.getDataFormat(), cellStyle.getDataFormatString());
                } else {
                    stringValue = m_dataFormatter.formatCellValue(cell);
                }
            }
            return stringValue;
        }
    }

}
