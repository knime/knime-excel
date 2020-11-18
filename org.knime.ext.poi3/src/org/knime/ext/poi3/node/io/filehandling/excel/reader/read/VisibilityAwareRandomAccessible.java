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
 *   Nov 16, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import org.knime.filehandling.core.node.table.reader.randomaccess.AbstractRandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;

/**
 * Implementation of {@link AbstractRandomAccessible} that is aware of its visibility.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 * @param <V> the type of values returned by this {@link RandomAccessible}
 */
public final class VisibilityAwareRandomAccessible<V> extends AbstractRandomAccessible<V> {

    private final RandomAccessible<V> m_randomAccessible;

    // true if the row is hidden in Excel
    private final boolean m_isHidden;

    /**
     * @param randomAccessible
     * @param isRowHidden
     */
    private VisibilityAwareRandomAccessible(final RandomAccessible<V> randomAccessible, final boolean isRowHidden) {
        m_randomAccessible = randomAccessible;
        m_isHidden = isRowHidden;
    }

    /**
     * Creates a new instance using <b>data</b> directly. Consequently, any changes to <b>data</b> will also change the
     * RandomAccessible. Use with caution.
     *
     * @param randomAccessible the underlying {@link RandomAccessible}
     * @param isHidden if whether is hidden or not
     * @return a new instance of VisibilityAwareRandomAccessible
     */
    public static <V> VisibilityAwareRandomAccessible<V> createUnsafe(final RandomAccessible<V> randomAccessible,
        final boolean isHidden) {
        return new VisibilityAwareRandomAccessible<>(randomAccessible, isHidden);
    }

    @Override
    public int size() {
        return m_randomAccessible.size();
    }

    @Override
    public V get(final int idx) {
        return m_randomAccessible.get(idx);
    }

    /**
     * @return the isRowHidden
     */
    boolean isHidden() {
        return m_isHidden;
    }

}
