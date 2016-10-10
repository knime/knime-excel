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
 *   29 Sep 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read2;

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Caches the seen date formats. <br/>
 * <b>NOT thread-safe</b>
 *
 * @author Gabor Bakos
 */
final class DateFormatCache {
    private final Set<String> m_dateFormat = new HashSet<>();
    private final Set<String> m_nonDateFormat = new HashSet<>();

    /**
     * Constructs the cache.
     */
    DateFormatCache() {
    }

    /**
     * @param formatCode
     * @param format
     * @return Checks whether it is a valid date format or not.
     */
    boolean isDateFormat(final int formatCode, final String format) {
        String key = key(formatCode, format);
        if (m_dateFormat.contains(key)) {
            return true;
        }
        if (m_nonDateFormat.contains(key)) {
            return false;
        }
        boolean ret = DateUtil.isADateFormat(formatCode, format);
        (ret ? m_dateFormat : m_nonDateFormat).add(key);
        return ret;
    }

    /**
     * @param formatCode
     * @param format
     * @return
     */
    private String key(final int formatCode, final String format) {
        return format == null ? Integer.toBinaryString(formatCode) + "\u0000\u0000"
            : Integer.toString(formatCode) + "\u0000" + format;
    }
}
