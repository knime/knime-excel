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

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * This class contains the KNIME type of a parsed Excel cell and its value stored as {@link Object}.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelCell {

    /**
     * The cell type of KNIME the {@link ExcelCell} maps to.
     */
    @SuppressWarnings("javadoc")
    public enum KNIMECellType {
            DOUBLE, LONG, INT, STRING, BOOLEAN, LOCAL_TIME, LOCAL_DATE, LOCAL_DATE_TIME;
    }

    private final KNIMECellType m_type;

    private final Object m_value;

    /**
     * Constructor.
     *
     * @param type the KNIME type of the cell
     * @param value the value of the cell
     */
    public ExcelCell(final KNIMECellType type, final Object value) {
        m_type = type;
        m_value = value;
    }

    /**
     * @return the type
     */
    public KNIMECellType getType() {
        return m_type;
    }

    @Override
    public String toString() {
        return getStringValue() + " (" + m_type + ")";
    }

    /**
     * @return the string value
     */
    public String getStringValue() {
        return String.valueOf(m_value);
    }

    /**
     * @return the double value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public double getDoubleValue() {
        if (m_type == KNIMECellType.DOUBLE) {
            return (double)m_value;
        }
        if (m_type == KNIMECellType.LONG) {
            return getLongValue();
        }
        if (m_type == KNIMECellType.INT) {
            return getIntValue();
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the long value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public long getLongValue() {
        if (m_type == KNIMECellType.LONG) {
            return (long)m_value;
        }
        if (m_type == KNIMECellType.INT) {
            return getIntValue();
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the int value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public int getIntValue() {
        if (m_type == KNIMECellType.INT) {
            return (int)m_value;
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the boolean value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public boolean getBooleanValue() {
        if (m_type == KNIMECellType.BOOLEAN) {
            return (boolean)m_value;
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the LocalTime value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public LocalTime getLocalTimeValue() {
        if (m_type == KNIMECellType.LOCAL_TIME) {
            return (LocalTime)m_value;
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the LocalDate value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public LocalDate getLocalDateValue() {
        if (m_type == KNIMECellType.LOCAL_DATE) {
            return (LocalDate)m_value;
        }
        throw unsupportedOperationException();
    }

    /**
     * @return the LocalDateTime value
     * @throws UnsupportedOperationException if the excel cell does not hold a compatible value
     */
    public LocalDateTime getLocalDateTimeValue() {
        if (m_type == KNIMECellType.LOCAL_DATE_TIME) {
            return (LocalDateTime)m_value;
        }
        if (m_type == KNIMECellType.LOCAL_DATE) {
            return getLocalDateValue().atStartOfDay();
        }
        throw unsupportedOperationException();
    }

    private UnsupportedOperationException unsupportedOperationException() {
        return new UnsupportedOperationException("Datatype '" + m_type + "' does not support this call.");
    }
}
