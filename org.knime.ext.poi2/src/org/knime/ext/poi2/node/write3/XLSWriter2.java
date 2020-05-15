/*
 * ------------------------------------------------------------------------
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
 * -------------------------------------------------------------------
 *
 * History
 *   Mar 15, 2007 (ohl): created
 */
package org.knime.ext.poi2.node.write3;

import java.awt.Dimension;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.data.BooleanValue;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.StringValue;
import org.knime.core.data.date.DateAndTimeValue;
import org.knime.core.data.image.png.PNGImageContent;
import org.knime.core.data.image.png.PNGImageValue;
import org.knime.core.data.time.duration.DurationValue;
import org.knime.core.data.time.localdate.LocalDateValue;
import org.knime.core.data.time.localdatetime.LocalDateTimeValue;
import org.knime.core.data.time.localtime.LocalTimeValue;
import org.knime.core.data.time.period.PeriodValue;
import org.knime.core.data.time.zoneddatetime.ZonedDateTimeValue;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;

/**
 *
 * @author ohl, University of Konstanz
 */
class XLSWriter2 {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSWriter2.class);

    /**
     * Excel (Ver. 2003) can handle datasheets up to 64k x 256 cells!
     */
    private static final int MAX_NUM_OF_ROWS = 65536;

    private static final int MAX_NUM_OF_COLS = 256;

    /**
     * Estimated value for pixel to value conversion.
     */
    private static final double PIXELS_TO_POINTS = 0.76;

    /**
     * 256 is one full character, one character has 7 pixels.
     */
    private static final double PIXELS_TO_CHARACTERS = 256 / 7;

    private final XLSWriter2Settings m_settings;

    private final Path m_destination;

    private final XLSNodeType m_type;

    /**
     * Creates a new writer with the specified settings.
     *
     * @param destination the destination to which the workbook will be written
     * @param type The Excel node type (writer or appender).
     * @param settings the settings.
     * @param choice the file system choice
     */
    XLSWriter2(final Path destination, final XLSNodeType type, final XLSWriter2Settings settings) {
        if (settings == null) {
            throw new IllegalArgumentException("Can't operate with null settings!");
        }
        m_destination = destination;
        m_type = type;
        m_settings = settings;
    }

    /**
     * Writes <code>table</code> with current settings.
     *
     * @param table the table to write to the file
     * @param spec {@link DataTableSpec} of {@code table}.
     * @param numOfRows
     * @param exec an execution monitor where to check for canceled status and report progress to. (In case of
     *            cancellation, the file will be deleted.)
     * @throws IOException if any related I/O error occurs
     * @throws CanceledExecutionException if execution in <code>exec</code> has been canceled
     * @throws InvalidFormatException if the existing file is not of the correct format
     * @throws URISyntaxException if the destination URL is invalid
     * @throws InvalidSettingsException
     * @throws NullPointerException if table is <code>null</code>
     */
    void write(final Iterable<DataRow> table, final DataTableSpec spec, final int numOfRows,
        final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException, InvalidFormatException, URISyntaxException, InvalidSettingsException {

        Workbook wb = null;
        final boolean isXLSX = m_destination.toString().toLowerCase().endsWith("xlsx");
        try {
            wb = readOrCreateWorkbook(isXLSX);

            int sheetIdx = 0; // in case the table doesn't fit in one sheet
            String sheetName = m_settings.getSheetname();
            if ((sheetName == null) || (sheetName.trim().length() == 0)) {
                sheetName = spec.getName();
            }
            // max sheetname length is 32 incl. added running index. We cut it to 25
            if (sheetName.length() > 25) {
                sheetName = sheetName.substring(0, 22) + "...";
            }
            // replace characters like \ / * ? [ ] etc.
            sheetName = replaceInvalidChars(sheetName);

            if (!m_settings.getDoNotOverwriteSheet()) {
                final Sheet oldSheet = wb.getSheet(sheetName);
                if (oldSheet != null) {
                    wb.removeSheetAt(wb.getSheetIndex(oldSheet));
                }
            }

            Sheet sheet = wb.createSheet(sheetName);
            if (m_settings.getAutosize() && (sheet instanceof SXSSFSheet)) {
                ((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
            }
            sheet.getPrintSetup().setLandscape(m_settings.getLandscape());
            sheet.getPrintSetup().setPaperSize(m_settings.getPaperSize());
            Drawing<?> drawing = null;
            final CreationHelper helper = wb.getCreationHelper();
            final Map<String, CellStyle> dateStyles = new HashMap<>();

            int numOfCols = spec.getNumColumns();
            final int rowHdrIncr = m_settings.writeRowID() ? 1 : 0;

            if (!isXLSX && ((numOfCols + rowHdrIncr) > MAX_NUM_OF_COLS)) {
                LOGGER.warn("The table to write has too many columns! Can't put more than " + MAX_NUM_OF_COLS
                    + " columns in one sheet. Truncating columns " + (MAX_NUM_OF_COLS + 1) + " to " + numOfCols);
                numOfCols = MAX_NUM_OF_COLS - rowHdrIncr;
            }

            int rowIdx = 0; // the index of the row in the XLsheet
            short colIdx = 0; // the index of the cell in the XLsheet

            // write column names
            if (m_settings.writeColHeader()) {

                // Create a new row in the sheet
                final Row hdrRow = sheet.createRow(rowIdx++);

                if (m_settings.writeRowID()) {
                    hdrRow.createCell(colIdx++).setCellValue("row ID");
                }
                for (int c = 0; c < numOfCols; c++) {
                    final String cName = spec.getColumnSpec(c).getName();
                    hdrRow.createCell(colIdx++).setCellValue(cName);
                }

            } // end of if write column names

            // Guess 80% of the job is generating the sheet, 20% is writing it out
            final ExecutionMonitor e = exec.createSubProgress(0.8);

            // Get the maximum sizes of rows and columns containing images
            final Map<Integer, Integer> columnWidth = new HashMap<Integer, Integer>();
            final Map<Integer, Integer> rowHeight = new HashMap<Integer, Integer>();
            final int colHdrIncr = m_settings.writeColHeader() ? 1 : 0;
            if (spec.containsCompatibleType(PNGImageValue.class)) {
                int i = 0;
                for (final DataRow row : table) {
                    for (int j = 0; j < numOfCols; j++) {
                        final DataCell cell = row.getCell(j);
                        if (!cell.isMissing() && cell.getType().isCompatible(PNGImageValue.class)) {
                            final Dimension dimension = ((PNGImageValue)cell).getImageContent().getPreferredSize();
                            final Integer currentHeight = rowHeight.get(i + colHdrIncr);
                            final Integer currentWidth = columnWidth.get(j + rowHdrIncr);
                            if ((currentHeight == null) || (currentHeight < dimension.height)) {
                                rowHeight.put(i + colHdrIncr, dimension.height);
                            }
                            if ((currentWidth == null) || (currentWidth < dimension.width)) {
                                columnWidth.put(j + rowHdrIncr, dimension.width);
                            }
                        }
                    }
                    i++;
                }
            }

            // write each row of the data
            int rowCnt = 0;
            for (final DataRow tableRow : table) {

                colIdx = 0;

                // create a new sheet if the old one is full
                if (!isXLSX && (rowIdx >= MAX_NUM_OF_ROWS)) {
                    sheetIdx++;
                    sheet = wb.createSheet(sheetName + "(" + sheetIdx + ")");
                    rowIdx = 0;
                    LOGGER.info("Creating additional sheet to store entire table. Additional sheet name: " + sheetName
                        + "(" + sheetIdx + ")");
                }

                // set the progress
                final String rowID = tableRow.getKey().getString();
                String msg;
                if (numOfRows <= 0) {
                    msg = "Writing row " + (rowCnt + 1) + " (\"" + rowID + "\")";
                } else {
                    msg = "Writing row " + (rowCnt + 1) + " (\"" + rowID + "\") of " + numOfRows;
                    e.setProgress(rowCnt / (double)numOfRows, msg);
                }
                // Check if execution was canceled !
                exec.checkCanceled();

                // Create a new row in the sheet
                final Row sheetRow = sheet.createRow(rowIdx++);

                // add the row id
                if (m_settings.writeRowID()) {
                    sheetRow.createCell(colIdx++).setCellValue(rowID);
                }
                // now add all data cells
                for (int c = 0; c < numOfCols; c++) {

                    final DataCell colValue = tableRow.getCell(c);

                    if (colValue.isMissing()) {
                        if (m_settings.getWriteMissingValues()) {
                            sheetRow.createCell(colIdx).setCellValue(m_settings.getMissingPattern());
                        }
                    } else {
                        final Cell sheetCell = sheetRow.createCell(colIdx);
                        if (colValue.getType().isCompatible(DoubleValue.class)) {
                            final double val = ((DoubleValue)colValue).getDoubleValue();
                            // Bug 6252 infinity is not supported by POI >= 3.9, we write a replacement formula instead
                            if (Double.isInfinite(val)) {
                                if (val > 0) {
                                    sheetCell.setCellFormula("1/0");
                                } else {
                                    sheetCell.setCellFormula("-1/0");
                                }
                            }
                            sheetCell.setCellValue(val);
                        } else if (colValue.getType().isCompatible(BooleanValue.class)) {
                            final boolean val = ((BooleanValue)colValue).getBooleanValue();
                            sheetCell.setCellValue(val);
                        } else if (colValue.getType().isCompatible(DateAndTimeValue.class)) {
                            final DateAndTimeValue dateAndTime = (DateAndTimeValue)colValue;
                            final Calendar val = dateAndTime.getUTCCalendarClone();
                            sheetCell.setCellValue(val);
                            String format = "";
                            if (dateAndTime.hasDate()) {
                                format += "yyyy-mm-dd";
                            }
                            if (dateAndTime.hasDate() && dateAndTime.hasTime()) {
                                format += "T";
                            }
                            if (dateAndTime.hasTime()) {
                                format += "hh:mm:ss";
                            }
                            if (dateAndTime.hasTime() && dateAndTime.hasMillis()) {
                                format += ".";
                            }
                            if (dateAndTime.hasMillis()) {
                                format += "000";
                            }
                            sheetCell.setCellStyle(dateOrTimeStyle(wb, helper, dateStyles, format));
                        } else if (colValue.getType().isCompatible(LocalDateValue.class)) {
                            final LocalDateValue dateValue = (LocalDateValue)colValue;
                            sheetCell.setCellStyle(dateOrTimeStyle(wb, helper, dateStyles, "yyyy-mm-dd"));
                            sheetCell.setCellValue(DateUtil.getExcelDate(new Date(dateValue.getLocalDate().atTime(0, 0)
                                .atZone(ZoneId.systemDefault()).toInstant().toEpochMilli())));
                        } else if (colValue.getType().isCompatible(PeriodValue.class)) {
                            final PeriodValue periodValue = (PeriodValue)colValue;
                            sheetCell.setCellValue(periodValue.getPeriod().toString());
                        } else if (colValue.getType().isCompatible(LocalDateTimeValue.class)) {
                            final LocalDateTimeValue dtValue = (LocalDateTimeValue)colValue;
                            sheetCell.setCellStyle(dateOrTimeStyle(wb, helper, dateStyles, "yyyy-mm-dd hh:mm:ss"));
                            sheetCell.setCellValue(DateUtil.getExcelDate(new Date(
                                dtValue.getLocalDateTime().atZone(ZoneId.systemDefault()).toInstant().toEpochMilli())));
                        } else if (colValue.getType().isCompatible(LocalTimeValue.class)) {
                            final LocalTimeValue tValue = (LocalTimeValue)colValue;
                            sheetCell.setCellStyle(dateOrTimeStyle(wb, helper, dateStyles, "hh:mm:ss;@"));
                            sheetCell.setCellValue(
                                DateUtil.getExcelDate(new Date(LocalDate.of(1900, 1, 1).atTime(tValue.getLocalTime())
                                    .atZone(ZoneId.systemDefault()).toInstant().toEpochMilli())) % 1);
                        } else if (colValue.getType().isCompatible(ZonedDateTimeValue.class)) {
                            final ZonedDateTimeValue zdtValue = (ZonedDateTimeValue)colValue;
                            sheetCell.setCellValue(zdtValue.toString());
                        } else if (colValue.getType().isCompatible(DurationValue.class)) {
                            final DurationValue durationValue = (DurationValue)colValue;
                            sheetCell.setCellValue(durationValue.getDuration().toString());
                        } else if (colValue.getType().isCompatible(PNGImageValue.class)) {
                            final PNGImageContent image = ((PNGImageValue)colValue).getImageContent();
                            final Dimension dimension = image.getPreferredSize();
                            final byte[] bytes = image.getByteArray();
                            final int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                            final double rowRatio =
                                (double)dimension.height / rowHeight.get(new Integer(sheetRow.getRowNum()));
                            final double columnRatio = (double)dimension.width / columnWidth.get(new Integer(colIdx));
                            final ClientAnchor anchor =
                                createAnchor(sheetRow.getRowNum(), colIdx, rowRatio, columnRatio, helper);
                            if (drawing == null) {
                                drawing = sheet.createDrawingPatriarch();
                            }
                            drawing.createPicture(anchor, pictureIdx);
                        } else if (colValue.getType().isCompatible(StringValue.class)) {
                            final String val = ((StringValue)colValue).getStringValue();
                            sheetCell.setCellValue(val);
                        } else {
                            final String val = colValue.toString();
                            sheetCell.setCellValue(val);
                        }
                    }

                    colIdx++;
                }
                if (rowHeight.containsKey(sheetRow.getRowNum())) {
                    sheetRow.setHeightInPoints(pixelsToPoints(rowHeight.get(sheetRow.getRowNum())));
                }

                rowCnt++;
            } // end of for all rows in table
            for (final Integer columnNumber : columnWidth.keySet()) {
                sheet.setColumnWidth(columnNumber, pixelsToCharacters(columnWidth.get(columnNumber)));
            }

            if (m_settings.getAutosize()) {
                for (int i = 0; i < numOfCols; i++) {
                    sheet.autoSizeColumn(i);
                }
            }

            if (m_settings.evaluateFormula()) {
                final FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                evaluator.evaluateAll();
            }

            writeWorkbook(wb);
        } finally {
            if (wb instanceof SXSSFWorkbook) {
                final SXSSFWorkbook sxssfWorkbook = (SXSSFWorkbook)wb;
                sxssfWorkbook.dispose();
            }
            if (wb != null) {
                wb.close();
            }
        }
    }

    private void writeWorkbook(final Workbook wb) throws IOException, InvalidSettingsException {

        final Set<OpenOption> openOptions = new HashSet<>();
        openOptions.add(StandardOpenOption.WRITE);

        if (m_type == XLSNodeType.WRITER) {
            if (m_settings.getOverwriteOK()) {
                // creates a file if it does not already exist, or truncates an existing file
                openOptions.add(StandardOpenOption.CREATE);
                openOptions.add(StandardOpenOption.TRUNCATE_EXISTING);
            } else {
                // fails with FileAlreadyExistsException if it already exists
                openOptions.add(StandardOpenOption.CREATE_NEW);
            }
        } else { // appender
            openOptions.add(StandardOpenOption.CREATE);
            openOptions.add(StandardOpenOption.TRUNCATE_EXISTING);
        }

        // Write the output to a file
        try (final OutputStream os =
            new BufferedOutputStream(Files.newOutputStream(m_destination, openOptions.toArray(new OpenOption[0])))) {
            wb.write(os);
            os.flush();
        } catch (FileAlreadyExistsException e) {
            throw new InvalidSettingsException(
                "Output file '" + m_destination + "' exists and must not be overwritten due to user settings");
        } // other IO Exceptions while writing will get thrown
    }

    private Workbook readOrCreateWorkbook(final boolean isXLSX) throws IOException, InvalidFormatException, InvalidSettingsException {
        Workbook wb;

        if (m_type == XLSNodeType.APPENDER) {
            try (InputStream inStream = Files.newInputStream(m_destination)) {
                wb = createWorkbookFromStream(isXLSX, inStream);
            } catch (NoSuchFileException e) {
                if (m_settings.getFileMustExist()) {
                    throw new InvalidSettingsException("Output location '" + m_destination + "' does not exist, append not possible");
                } else {
                    // the file does not exist, hence we start with a fresh workbook
                    wb = createNewWorkbook(isXLSX);
                }
            } // other IO Exceptions while reading will get thrown
        } else {
            // for the writer node we always start with a new workbook
            wb = createNewWorkbook(isXLSX);
        }

        return wb;
    }

    /**
     * @param isXLSX
     * @return
     */
    private Workbook createNewWorkbook(final boolean isXLSX) {
        final Workbook wb;
        if (isXLSX) {
            wb = createSxssfWorkbook();
            ((SXSSFWorkbook)wb).setCompressTempFiles(true);
        } else {
            wb = new HSSFWorkbook();
        }
        return wb;
    }

    /**
     * @param wb Workbook
     * @param helper creation helper
     * @param dateStyles used date styles
     * @param format format to use
     * @return {@link CellStyle} of the date.
     */
    private static CellStyle dateOrTimeStyle(final Workbook wb, final CreationHelper helper,
        final Map<String, CellStyle> dateStyles, final String format) {
        CellStyle dateStyle;
        if (dateStyles.containsKey(format)) {
            dateStyle = dateStyles.get(format);
        } else {
            dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(helper.createDataFormat().getFormat(format));
            dateStyles.put(format, dateStyle);
        }
        return dateStyle;
    }

    /**
     * @param isXLSX Is the stream an xlsx file?
     * @param inStream An {@link InputStream} of an xls/xlsx file.
     * @return The {@link Workbook} for the {@code inStream}.
     * @throws IOException Problem reading.
     * @throws InvalidFormatException Wrong format.
     */
    private Workbook createWorkbookFromStream(final boolean isXLSX, final InputStream inStream)
        throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(inStream);
        if (isXLSX && !m_settings.evaluateFormula() && (wb instanceof XSSFWorkbook)) {
            wb = new SXSSFWorkbook((XSSFWorkbook)wb);
        }
        return wb;
    }

    /**
     * @return A streaming version of xlsx workbook.
     */
    private static SXSSFWorkbook createSxssfWorkbook() {
        return new SXSSFWorkbook(100);
    }

    private static int pixelsToCharacters(final int pixels) {
        // added 2 margin pixels per side and 1 for the gridline
        return (int)((pixels + 5) * PIXELS_TO_CHARACTERS);
    }

    private static int pixelsToPoints(final int pixels) {
        return (int)(pixels * PIXELS_TO_POINTS);
    }

    /**
     * Creates an anchor that fits to a certain cell.
     *
     *
     * @param row Row number of the cell
     * @param column Column number of the cell
     * @param helper Helper that creates the anchor object
     * @return Initialized anchor object
     */
    private static ClientAnchor createAnchor(final int row, final int column, final double rowRatio,
        final double columnRatio, final CreationHelper helper) {
        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
        anchor.setCol1(column);
        anchor.setRow1(row);
        anchor.setCol2(column);
        anchor.setRow2(row);
        // need to make the image smaller than the cell so it is sortable
        // create a margin of 1 on each side
        if (anchor instanceof XSSFClientAnchor) {
            // format is XSSF
            // full cell size is 1023 width and 255 height
            // each multiplied by XSSFShape.EMU_PER_PIXEL
            anchor.setDx1(1);
            anchor.setDy1(1);
            anchor.setDx2((int)(1023 * XSSFShape.EMU_PER_PIXEL * columnRatio) - 1);
            anchor.setDy2((int)(255 * XSSFShape.EMU_PER_PIXEL * rowRatio) - 1);
        } else {
            // format is HSSF
            // full cell size is 1023 width and 255 height
            anchor.setDx1(1);
            anchor.setDy1(1);
            anchor.setDx2((int)(1022 * columnRatio));
            anchor.setDy2((int)(254 * rowRatio));
        }
        return anchor;
    }

    /**
     * Replaces characters that are illegal in sheet names. These are \/:*?"<>|[].
     *
     * @param name the name to clean
     * @return returns the name with all of the above characters replaced by an underscore.
     */
    private static String replaceInvalidChars(final String name) {
        final StringBuilder result = new StringBuilder();
        final int l = name.length();
        for (int i = 0; i < l; i++) {
            final char c = name.charAt(i);
            if ((c == '\\') || (c == '/') || (c == ':') || (c == '*') || (c == '?') || (c == '"') || (c == '<')
                || (c == '>') || (c == '|') || (c == '[') || (c == ']')) {
                result.append('_');
            } else {
                result.append(c);
            }
        }
        return result.toString();
    }

}
