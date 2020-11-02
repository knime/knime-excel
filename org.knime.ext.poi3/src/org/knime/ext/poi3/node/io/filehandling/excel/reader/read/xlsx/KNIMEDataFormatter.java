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
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xlsx;

import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;

/**
 * Extension of {@link DataFormatter} that creates numeric and date&time {@link ExcelCell}s directly instead of
 * formatting numeric values. This prevents unnecessary parsing of formatted Strings.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class KNIMEDataFormatter extends DataFormatter {

    /**
     * The {@code use1904Windowing} passed into {@link #formatRawCellContents(double, int, String, boolean)} is always
     * false, we need our own flag for it.
     */
    private final boolean m_use1904Windowing;

    private final boolean m_use15DigitsPrecision;

    private ExcelCell m_excelCell;

    KNIMEDataFormatter(final boolean use1904Windowing, final boolean use15DigitsPrecision) {
        m_use1904Windowing = use1904Windowing;
        m_use15DigitsPrecision = use15DigitsPrecision;
    }

    /**
     * Does not do any formatting and returns always {@code null}. This function is overridden in order to pickup the
     * double value and create the {@code ExcelCell} directly. It also allows us to directly check whether the cell is a
     * date or numeric. Hence, we save a lot of parsing etc. The super function is only called when necessary to save
     * execution time.
     */
    @Override
    public String formatRawCellContents(final double value, final int formatIndex, final String formatString,
        final boolean use1904Windowing) {
        if (DateUtil.isADateFormat(formatIndex, formatString)) {
            final LocalDateTime ldt = DateUtil.getLocalDateTime(value, m_use1904Windowing);
            if (ldt == null) {
                // this can usually not happen if the file has been created with Excel software; however,
                // e.g., LibreOffice or the KNIME Excel Writer allows to create dates before 01-01-1900 which
                // leads to this case; we will output the underlying double value in this case
                m_excelCell = new ExcelCell(KNIMECellType.DOUBLE, value);
            } else {
                m_excelCell = ExcelCellUtils.createDateTimeExcelCell(ldt, m_use1904Windowing);
            }
        } else {
            final ExcelCell excelCell =
                m_use15DigitsPrecision ? ExcelCellUtils.createNumericExcelCellUsing15DigitsPrecision(value)
                    : ExcelCellUtils.createNumericExcelCell(value);
            // we might have a Boolean if the value is 0 or 1
            if (ExcelCellUtils.canBeBoolean(excelCell)) {
                final String formattedString =
                    super.formatRawCellContents(value, formatIndex, formatString, use1904Windowing);
                m_excelCell = ExcelCellUtils.getIntOrBooleanCell(excelCell, formattedString);
            } else {
                m_excelCell = excelCell;
            }
        }
        return null;
    }

    /**
     * Returns the last created {@link ExcelCell} or {@code null} if there is none. Also, sets the {@link ExcelCell} to
     * {@code null}.
     */
    ExcelCell getAndResetExcelCell() {
        final ExcelCell excelCell = m_excelCell;
        m_excelCell = null;
        return excelCell;
    }

}
