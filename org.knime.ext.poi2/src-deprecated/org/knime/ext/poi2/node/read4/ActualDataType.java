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
 *   23 Aug 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read4;

/**
 * The data type as seen from the file.
 *
 * @author Gabor Bakos
 */
@Deprecated
public enum ActualDataType {
    /** The missing value */
    MISSING,
    /** Erroneous value */
    ERROR,
    /** Boolean value */
    BOOLEAN,
    /** Number with unknown precision. */
    NUMBER,
    /** Integer number. */
    NUMBER_INT,
    /** Floating point number */
    NUMBER_DOUBLE,
    /** Text */
    STRING,
    /** Date/time */
    DATE,
    /** Formula with missing value as a result. (Probably cannot happen.) */
    MISSING_FORMULA,
    /** Formula with error result. */
    ERROR_FORMULA,
    /** Formula with boolean result. */
    BOOLEAN_FORMULA,
    /** Formula with numeric result (of unknown precision). */
    NUMBER_FORMULA,
    /** Formula with integer result. */
    NUMBER_INT_FORMULA,
    /** Formula with floating point number result. */
    NUMBER_DOUBLE_FORMULA,
    /** Formula with text as a result. */
    STRING_FORMULA,
    /** Formula with date/time as a result. */
    DATE_FORMULA;

    /**
     * @param type An {@link ActualDataType}.
     * @return Whether the type is a possible formula or not.
     */
    public static boolean isFormula(final ActualDataType type) {
        switch (type) {
            case BOOLEAN:
            case DATE:
            case ERROR:
            case MISSING:
            case NUMBER:
            case NUMBER_DOUBLE:
            case NUMBER_INT:
            case STRING:
                return false;
            case BOOLEAN_FORMULA:
            case DATE_FORMULA:
            case ERROR_FORMULA:
            case MISSING_FORMULA:
            case NUMBER_FORMULA:
            case NUMBER_DOUBLE_FORMULA:
            case NUMBER_INT_FORMULA:
            case STRING_FORMULA:
                return true;
            default:
                throw new IllegalStateException("Unknown: " + type);
        }
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is a {@link #DATE} or {@link #DATE_FORMULA}.
     */
    public static boolean isDate(final ActualDataType type) {
        return type == ActualDataType.DATE || type == ActualDataType.DATE_FORMULA;
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is a {@link #BOOLEAN} or {@link #BOOLEAN_FORMULA}.
     */
    public static boolean isBoolean(final ActualDataType type) {
        return type == ActualDataType.BOOLEAN || type == ActualDataType.BOOLEAN_FORMULA;
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is a {@link #STRING} or {@link #STRING_FORMULA}.
     */
    public static boolean isString(final ActualDataType type) {
        return type == ActualDataType.STRING || type == ActualDataType.STRING_FORMULA;
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is a {@link #NUMBER}, {@link #NUMBER_INT}, {@link #NUMBER_DOUBLE},
     *         {@value #NUMBER_FORMULA}, {@link #NUMBER_INT_FORMULA} or {@link #NUMBER_DOUBLE_FORMULA}.
     */
    public static boolean isNumber(final ActualDataType type) {
        return type == ActualDataType.NUMBER || type == ActualDataType.NUMBER_DOUBLE
            || type == ActualDataType.NUMBER_INT || type == ActualDataType.NUMBER_FORMULA
            || type == ActualDataType.NUMBER_DOUBLE_FORMULA || type == ActualDataType.NUMBER_INT_FORMULA;
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is a {@link #MISSING} or {@link #MISSING_FORMULA}.
     */
    public static boolean isMissing(final ActualDataType type) {
        return type == ActualDataType.MISSING || type == ActualDataType.MISSING_FORMULA;
    }

    /**
     * @param type An {@link ActualDataType}.
     * @return {@code true} iff it is an {@link #ERROR} or {@link #ERROR_FORMULA}.
     */
    public static boolean isError(final ActualDataType type) {
        return type == ActualDataType.ERROR || type == ActualDataType.ERROR_FORMULA;
    }
}
