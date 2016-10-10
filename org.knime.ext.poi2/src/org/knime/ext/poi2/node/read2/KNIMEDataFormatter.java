/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
package org.knime.ext.poi2.node.read2;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * KNIME specific data formatter which encodes the type of the value in the first character of the formatted value.
 * It prepends invaliid char-sequences to the output to provide type information.
 *
 * @author Gabor Bakos
 */
class KNIMEDataFormatter extends DataFormatter {
    private final SimpleDateFormat m_dateAndTimeFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss"),
            m_dateFormat = new SimpleDateFormat("yyyy/MM/dd");

    //Intentionally invalid UFT-16 encoding to avoid collision with valid UTF-16 Strings:
    //https://en.wikipedia.org/wiki/UTF-16
    static final String DATE_PREFIX = "\uDC37\uDC37", NUMBER_PREFIX = "\uDC73\uDC73", BOOL_PREFIX = "\uDC45\uDC45";

    private static final String TRUE = Boolean.toString(true).toUpperCase(),
            FALSE = Boolean.toString(false).toUpperCase();

    private static final int MILLISEC_PER_DAY = 24*3600;

    /**
     * The date format to be produced.
     */
    public enum DateFormat {
        /** ISO-9801-like date formats */
        Standardized,
        /** Date format specified in xlsx. */
        ExcelFormat
    }

    private static final DateFormat DEFAULT_STANDARDIZE_DATE = DateFormat.Standardized;

    private final DateFormat m_standardizeDate;

    private String m_lastOriginalFormattedValue;

    private final DateFormatCache m_dateFormatCache = new DateFormatCache();

    /**
     * Consructs a data formatter with standardized date formatting and with default Locale.
     */
    KNIMEDataFormatter() {
        this(DEFAULT_STANDARDIZE_DATE);
    }

    /**
     * Creates a {@link KNIMEDataFormatter} with possibly non-standardizing {@link DateFormat}.
     *
     * @param df A custom {@link DateFormat}.
     */
    KNIMEDataFormatter(final DateFormat df) {
        m_standardizeDate = df;
    }

    /**
     * Using standardized {@link DateFormat}.
     *
     * @param locale The locale of the formatting to use.
     */
    KNIMEDataFormatter(final Locale locale) {
        this(locale, DEFAULT_STANDARDIZE_DATE);
    }

    /**
     * @param locale The {@link Locale} to use.
     * @param dateFormat {@link DateFormat} to use.
     */
    protected KNIMEDataFormatter(final Locale locale, final DateFormat dateFormat) {
        super(locale);
        m_standardizeDate = dateFormat;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String formatRawCellContents(final double value, final int formatIndex, final String formatString,
        final boolean use1904Windowing) {
        String orig = super.formatRawCellContents(value, formatIndex, formatString, use1904Windowing);
        if (m_dateFormatCache.isDateFormat(formatIndex, formatString)) {
            if (DateUtil.isValidExcelDate(value)) {
                m_lastOriginalFormattedValue = DATE_PREFIX + orig;
                switch (m_standardizeDate) {
                    case Standardized:
                        Date date = getRoundedDate(value, use1904Windowing);
                        if ((long)value == value) {
                            return DATE_PREFIX + m_dateFormat.format(date);
                        }
                        return DATE_PREFIX + m_dateAndTimeFormat.format(date);
                    case ExcelFormat:
                        return m_lastOriginalFormattedValue;
                }
            }
        }
        if (orig.startsWith(TRUE) || orig.startsWith(FALSE)) {
            m_lastOriginalFormattedValue = BOOL_PREFIX + orig;
            return m_lastOriginalFormattedValue;
        }
        m_lastOriginalFormattedValue = NUMBER_PREFIX + orig;
        return NUMBER_PREFIX + value;
    }

    private static Date getRoundedDate(final double date, final boolean use1904Windowing) {
        int wholeDays = (int)Math.floor(date);
        double ms = date - wholeDays;
        int millisecondsInDay = (int)Math.round(MILLISEC_PER_DAY * ms) * 1000;

        Calendar calendar = new GregorianCalendar();
        DateUtil.setCalendar(calendar, wholeDays, millisecondsInDay, use1904Windowing, false);
        return calendar.getTime();
    }

    /**
     * @return the last value with the original formatting with the type prefix.
     */
    public String getLastOriginalFormattedValue() {
        return m_lastOriginalFormattedValue;
    }
}
