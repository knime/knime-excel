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
 *   Nov 6, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.cell;

import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.knime.core.data.BooleanValue;
import org.knime.core.data.DataCell;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.IntValue;
import org.knime.core.data.LongValue;
import org.knime.core.data.StringValue;
import org.knime.core.data.date.DateAndTimeValue;
import org.knime.core.data.time.duration.DurationValue;
import org.knime.core.data.time.localdate.LocalDateValue;
import org.knime.core.data.time.localdatetime.LocalDateTimeValue;
import org.knime.core.data.time.localtime.LocalTimeValue;
import org.knime.core.data.time.period.PeriodValue;
import org.knime.core.data.time.zoneddatetime.ZonedDateTimeValue;
import org.knime.filehandling.core.data.location.FSLocationValue;

/**
 * Convenience {@link ExcelCellWriter} writer that takes care of missing cells. Any extending class solely has to write
 * the {@link DataCell}'s value to the {@link Cell}.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
abstract class AbstractCellWriter implements ExcelCellWriter {

    private final String m_missingValPattern;

    AbstractCellWriter(final String missingValPattern) {
        m_missingValPattern = missingValPattern;
    }

    @Override
    public void write(final DataCell dataCell, final Function<CellType, Cell> cellCreator) throws IOException {
        if (dataCell.isMissing()) {
            if (m_missingValPattern != null) {
                cellCreator.apply(CellType.STRING).setCellValue(m_missingValPattern);
            } else {
                cellCreator.apply(CellType.BLANK);
            }
        } else {
            fill(dataCell, cellCreator.apply(getCellType()));
        }
    }

    abstract void fill(final DataCell dataCell, final Cell cell) throws IOException;

    abstract CellType getCellType();

}

final class BooleanCellWriter extends AbstractCellWriter {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    BooleanCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((BooleanValue)dataCell).getBooleanValue());
    }

    @Override
    public CellType getCellType() {
        return CellType.BOOLEAN;
    }

}

abstract class NumericCellWriters extends AbstractCellWriter {

    /**
     * Constructor.
     *
     * @param missingValPattern
     */
    NumericCellWriters(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    public final CellType getCellType() {
        return CellType.NUMERIC;
    }

}

final class IntCellWriter extends NumericCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    IntCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((IntValue)dataCell).getIntValue());
    }

}

final class LongCellWriter extends NumericCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    LongCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((LongValue)dataCell).getLongValue());
    }

}

final class DoubleCellWriter extends NumericCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    DoubleCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) throws IOException {
        final double val = ((DoubleValue)dataCell).getDoubleValue();
        if (Double.isInfinite(val)) {
            throw new IOException("Excel does not support inifinity");
        }
        cell.setCellValue(val);
    }

}

abstract class StyledCellWriters extends NumericCellWriters {

    private final CellStyleManager m_cellStyleManager;

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     * @param cellStyleManager the manager to create the appropriate cell style
     */
    StyledCellWriters(final String missingValPattern, final CellStyleManager cellStyleManager) {
        super(missingValPattern);
        m_cellStyleManager = cellStyleManager;
    }

    /**
     * Sets the style of the cell.
     *
     * @param cell the cell whose style has to be set
     * @param format the style format
     */
    void setCellStyle(final Cell cell, final String format) {
        cell.setCellStyle(m_cellStyleManager.getCellStyle(format));
    }
}

final class DateAndTimeCellWriter extends StyledCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     * @param cellStyleManager the manager to create the appropriate cell style
     */
    DateAndTimeCellWriter(final String missingValPattern, final CellStyleManager cellStyleManager) {
        super(missingValPattern, cellStyleManager);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        final DateAndTimeValue dateAndTime = (DateAndTimeValue)dataCell;
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
        setCellStyle(cell, format);
        cell.setCellValue(DateUtil.getExcelDate(dateAndTime.getUTCCalendarClone(), false));
    }

}

final class LocalDateCellWriter extends StyledCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     * @param cellStyleManager the manager to create the appropriate cell style
     */
    LocalDateCellWriter(final String missingValPattern, final CellStyleManager cellStyleManager) {
        super(missingValPattern, cellStyleManager);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        final LocalDateValue val = (LocalDateValue)dataCell;
        setCellStyle(cell, "yyyy-mm-dd");
        cell.setCellValue(DateUtil.getExcelDate(val.getLocalDate().atStartOfDay()));
    }

}

final class PeriodValueCellWriter extends NumericCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    PeriodValueCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((PeriodValue)dataCell).getPeriod().toString());
    }

}

final class LocalDateTimeCellWriter extends StyledCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     * @param cellStyleManager the manager to create the appropriate cell style
     */
    LocalDateTimeCellWriter(final String missingValPattern, final CellStyleManager cellStyleManager) {
        super(missingValPattern, cellStyleManager);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        final LocalDateTimeValue val = (LocalDateTimeValue)dataCell;
        setCellStyle(cell, "yyyy-mm-dd hh:mm:ss");
        cell.setCellValue(DateUtil.getExcelDate(val.getLocalDateTime()));
    }

}

final class LocalTimeCellWriter extends StyledCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     * @param cellStyleManager the manager to create the appropriate cell style
     */
    LocalTimeCellWriter(final String missingValPattern, final CellStyleManager cellStyleManager) {
        super(missingValPattern, cellStyleManager);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        final LocalTimeValue val = (LocalTimeValue)dataCell;
        setCellStyle(cell, "hh:mm:ss;@");
        cell.setCellValue(DateUtil.getExcelDate(LocalDateTime.of(LocalDate.of(1900, 1, 1), val.getLocalTime())) % 1);
    }

}

final class DurationCellWriter extends NumericCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    DurationCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((DurationValue)dataCell).getDuration().toString());
    }

}

abstract class StringCellWriters extends AbstractCellWriter {

    /**
     * Constructor.
     *
     * @param missingValPattern
     */
    StringCellWriters(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    public final CellType getCellType() {
        return CellType.STRING;
    }

}

final class ZonedDateTimeCellWriter extends StringCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    ZonedDateTimeCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((ZonedDateTimeValue)dataCell).getZonedDateTime().toString());
    }

}

final class FSLocationCellWriter extends StringCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    FSLocationCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((FSLocationValue)dataCell).getFSLocation().getPath());
    }

}

final class StringCellWriter extends StringCellWriters {

    /**
     * Constructor.
     *
     * @param missingValPattern the missing value pattern
     */
    StringCellWriter(final String missingValPattern) {
        super(missingValPattern);
    }

    @Override
    void fill(final DataCell dataCell, final Cell cell) {
        cell.setCellValue(((StringValue)dataCell).getStringValue());
    }

}
