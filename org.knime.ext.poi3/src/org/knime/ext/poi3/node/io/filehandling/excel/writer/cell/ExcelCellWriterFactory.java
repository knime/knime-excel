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
 *   Nov 12, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.cell;

import java.io.IOException;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.knime.core.data.BooleanValue;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataType;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.IntValue;
import org.knime.core.data.LongValue;
import org.knime.core.data.StringValue;
import org.knime.core.data.date.DateAndTimeValue;
import org.knime.core.data.image.png.PNGImageValue;
import org.knime.core.data.time.duration.DurationValue;
import org.knime.core.data.time.localdate.LocalDateValue;
import org.knime.core.data.time.localdatetime.LocalDateTimeValue;
import org.knime.core.data.time.localtime.LocalTimeValue;
import org.knime.core.data.time.period.PeriodValue;
import org.knime.core.data.time.zoneddatetime.ZonedDateTimeValue;
import org.knime.filehandling.core.data.location.FSLocationValue;

/**
 * Factory for {@link ExcelCellWriter}.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public class ExcelCellWriterFactory {

    private final CellStyleManager m_formatManager;

    private final String m_missingValPattern;

    private ExcelCellWriterFactory(final Workbook wb, final String missingValPattern) {
        m_formatManager = new CellStyleManager(wb);
        m_missingValPattern = missingValPattern;
    }

    /**
     * Creates the {@link ExcelCellWriterFactory}.
     *
     * @param wb the {@link Workbook}
     * @param missingValPattern the missing value pattern
     * @return the {@link ExcelCellWriterFactory}
     */
    public static ExcelCellWriterFactory createFactory(final Workbook wb, final String missingValPattern) {
        return new ExcelCellWriterFactory(wb, missingValPattern);
    }

    /**
     * Creates the {@link ExcelCellWriter} associated with the given {@link DataType}.
     *
     * @param type the {@link DataType}
     * @return the {@link ExcelCellWriter} associated with the given {@link DataType}
     */
    public ExcelCellWriter createCellWriter(final DataType type) { // NOSONAR 5 is not enough
        if (type.isCompatible(BooleanValue.class)) {
            return new BooleanCellWriter(m_missingValPattern);
        } else if (type.isCompatible(IntValue.class)) {
            return new IntCellWriter(m_missingValPattern);
        } else if (type.isCompatible(LongValue.class)) {
            return new LongCellWriter(m_missingValPattern);
        } else if (type.isCompatible(DoubleValue.class)) {
            return new DoubleCellWriter(m_missingValPattern);
        } else if (type.isCompatible(DateAndTimeValue.class)) {
            return new DateAndTimeCellWriter(m_missingValPattern, m_formatManager);
        } else if (type.isCompatible(LocalDateValue.class)) {
            return new LocalDateCellWriter(m_missingValPattern, m_formatManager);
        } else if (type.isCompatible(PeriodValue.class)) {
            return new PeriodValueCellWriter(m_missingValPattern);
        } else if (type.isCompatible(LocalDateTimeValue.class)) {
            return new LocalDateTimeCellWriter(m_missingValPattern, m_formatManager);
        } else if (type.isCompatible(LocalTimeValue.class)) {
            return new LocalTimeCellWriter(m_missingValPattern, m_formatManager);
        } else if (type.isCompatible(ZonedDateTimeValue.class)) {
            return new ZonedDateTimeCellWriter(m_missingValPattern);
        } else if (type.isCompatible(DurationValue.class)) {
            return new DurationCellWriter(m_missingValPattern);
        } else if (type.isCompatible(FSLocationValue.class)) {
            return new FSLocationCellWriter(m_missingValPattern);
        } else if (type.isCompatible(PNGImageValue.class)) {
            return new ExcelCellWriter() {

                @Override
                public void write(final DataCell dataCell, final Function<CellType, Cell> cellCreator)
                    throws IOException {
                    // Image are no regular cells and must be treated by the writer itself
                }
            };
        } else if (type.isCompatible(StringValue.class)) {
            return new StringCellWriter(m_missingValPattern);
        }
        throw new IllegalArgumentException(String.format(
            "Unsupported column '%s'. Please remove corresponding columns from the input table.", type.getName()));
    }

}
