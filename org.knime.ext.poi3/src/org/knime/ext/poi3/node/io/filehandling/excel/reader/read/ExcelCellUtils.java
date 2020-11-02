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
 *   Oct 27, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.time.LocalDateTime;

import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;

import com.google.common.math.DoubleMath;

/**
 * Utility class for {@link ExcelCell}s.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelCellUtils {

    private static final MathContext MATH_CONTEXT = new MathContext(15, RoundingMode.HALF_UP);

    private ExcelCellUtils() {
        // Hide constructor, utils class
    }

    private static final String TRUE = Boolean.toString(true).toUpperCase();

    private static final String FALSE = Boolean.toString(false).toUpperCase();

    /**
     * Creates the proper date&time cell, which can be of type {@link KNIMECellType#LOCAL_DATE_TIME},
     * {@link KNIMECellType#LOCAL_DATE}, or {@link KNIMECellType#LOCAL_TIME}.
     *
     * @param localDateTimeCellValue the {@link LocalDateTime}
     * @param use1904Windowing if 1904 windowing is used
     * @return the proper {@link ExcelCell} either of type LocalDate, LocalTime, or LocalDateTime
     */
    public static ExcelCell createDateTimeExcelCell(final LocalDateTime localDateTimeCellValue,
        final boolean use1904Windowing) {
        final ExcelCell excelCell;
        // Cells with only time but no date value have date set to 1899-12-31 or 1904-01-01 depending on if the date is
        // using 1904 windowing.
        if (localDateTimeCellValue.getYear() == (use1904Windowing ? 1904 : 1899)) {
            excelCell = new ExcelCell(KNIMECellType.LOCAL_TIME, localDateTimeCellValue.toLocalTime());
        } else {
            // We cannot know if only a date or date&time is set. Only-date values have time set
            // to 00:00 which could be midnight for date&time. The type mapping will take care of such cases and append
            // the 00:00 back to the LocalDate if necessary. Hence, we assume it is only a date here.
            if (isMidnight(localDateTimeCellValue)) {
                excelCell = new ExcelCell(KNIMECellType.LOCAL_DATE, localDateTimeCellValue.toLocalDate());
            } else {
                excelCell = new ExcelCell(KNIMECellType.LOCAL_DATE_TIME, localDateTimeCellValue);
            }
        }
        return excelCell;
    }

    private static boolean isMidnight(final LocalDateTime ldt) {
        return ldt.getHour() == 0 && ldt.getMinute() == 0 && ldt.getSecond() == 0 && ldt.getNano() == 0;
    }

    /**
     * Corrects potential floating point issues of the the passed double by rounding with 15 digits precision and checks
     * if it fits into int or long. Returns the proper {@link ExcelCell}.
     *
     * @param cellValue the numeric cell value
     * @return the proper {@link ExcelCell} either of type Double, Long, or Int
     */
    public static ExcelCell createNumericExcelCellUsing15DigitsPrecision(final double cellValue) {
        final ExcelCell excelCell;
        // we do not use BigDecimal#valueOf as it converts the double into String which is then parsed (expensive).
        // as we are rounding anyway afterwards, passing the double directly into the constructor is fine
        final BigDecimal bd = new BigDecimal(cellValue); // NOSONAR
        final BigDecimal roundedBD = bd.round(MATH_CONTEXT);
        final double doubleValue = roundedBD.doubleValue();
        // check if we actually have an int or long and create the proper ExcelCell
        if (DoubleMath.isMathematicalInteger(doubleValue)) {
            if (doubleValue <= Integer.MAX_VALUE && doubleValue >= Integer.MIN_VALUE) {
                excelCell = new ExcelCell(KNIMECellType.INT, roundedBD.intValue());
            } else if (doubleValue <= Long.MAX_VALUE && doubleValue >= Long.MIN_VALUE) {
                excelCell = new ExcelCell(KNIMECellType.LONG, roundedBD.longValue());
            } else {
                excelCell = new ExcelCell(KNIMECellType.DOUBLE, doubleValue);
            }
        } else {
            excelCell = new ExcelCell(KNIMECellType.DOUBLE, doubleValue);
        }
        return excelCell;
    }

    /**
     * Checks if the passed double fits into int or long. Returns the proper {@link ExcelCell}.
     *
     * @param cellValue the numeric cell value
     * @return the proper {@link ExcelCell} either of type Double, Long, or Int
     */
    public static ExcelCell createNumericExcelCell(final double cellValue) {
        final ExcelCell excelCell;
        // check if we actually have an int or long and create the proper ExcelCell
        if (DoubleMath.isMathematicalInteger(cellValue)) {
            if (cellValue <= Integer.MAX_VALUE && cellValue >= Integer.MIN_VALUE) {
                excelCell = new ExcelCell(KNIMECellType.INT, (int)cellValue);
            } else if (cellValue <= Long.MAX_VALUE && cellValue >= Long.MIN_VALUE) {
                excelCell = new ExcelCell(KNIMECellType.LONG, (long)cellValue);
            } else {
                excelCell = new ExcelCell(KNIMECellType.DOUBLE, cellValue);
            }
        } else {
            excelCell = new ExcelCell(KNIMECellType.DOUBLE, cellValue);
        }
        return excelCell;
    }

    /**
     * Checks if the passed {@link ExcelCell} could be of type boolean.
     *
     * @param excelCell the {@link ExcelCell}
     * @return true if it can be boolean
     */
    public static boolean canBeBoolean(final ExcelCell excelCell) {
        return excelCell.getType() == KNIMECellType.INT
            && (excelCell.getIntValue() == 0 || excelCell.getIntValue() == 1);
    }

    /**
     * Checks if the passed formatted String stands for a boolean and returns the proper {@link ExcelCell}. If not
     * boolean, the passed cell is returned.
     *
     * The case that it is a boolean is actually very rare and should only happen in files created with third party
     * software (I was somehow able to create such a file with LibreOffice).
     *
     * @param excelCell the integer {@link ExcelCell}
     * @param formattedString the formatted string
     * @return the proper boolean {@link ExcelCell} or the passed integer one
     */
    public static ExcelCell getIntOrBooleanCell(final ExcelCell excelCell, final String formattedString) {
        // #startsWith as for xlsx we get "TRUE+"/"FALSE+" and for xls "TRUE"/"FALSE"
        if (formattedString.startsWith(TRUE)) {
            return new ExcelCell(KNIMECellType.BOOLEAN, true);
        } else if (formattedString.startsWith(FALSE)) {
            return new ExcelCell(KNIMECellType.BOOLEAN, false);
        } else {
            return excelCell;
        }
    }
}
