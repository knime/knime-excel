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

import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.function.Supplier;

import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;
import org.knime.filehandling.core.node.table.reader.read.Read;
import org.knime.filehandling.core.node.table.reader.spec.ExtractColumnHeaderRead;
import org.knime.filehandling.core.util.CheckedExceptionSupplier;

/**
 * A {@link Read} that wraps another {@link Read} as an {@link ExtractColumnHeaderRead} that does nothing besides
 * passing the column headers.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class WrapperExtractColumnHeaderRead implements ExtractColumnHeaderRead<ExcelCell> {

    /** The underlying read. */
    private final Read<ExcelCell> m_read;

    private CheckedExceptionSupplier<Optional<RandomAccessible<ExcelCell>>, IOException> m_columnHeadersSupplier;

    /**
     * @param source the source {@link Read}
     * @param columnHeadersSupplier a {@link Supplier} that supplies the column headers and can throw
     *            {@link IOException}s
     */
    public WrapperExtractColumnHeaderRead(final Read<ExcelCell> source,
        final CheckedExceptionSupplier<Optional<RandomAccessible<ExcelCell>>, IOException> columnHeadersSupplier) {
        m_read = source;
        m_columnHeadersSupplier = columnHeadersSupplier;
    }

    @Override
    public RandomAccessible<ExcelCell> next() throws IOException {
        return m_read.next();
    }

    @Override
    public OptionalLong getMaxProgress() {
        return m_read.getMaxProgress();
    }

    @Override
    public long getProgress() {
        return m_read.getProgress();
    }

    @Override
    public void close() throws IOException {
        m_read.close();
    }

    @Override
    public Optional<RandomAccessible<ExcelCell>> getColumnHeaders() throws IOException {
        final Optional<RandomAccessible<ExcelCell>> headers = m_columnHeadersSupplier.get();
        if (headers.isPresent()) {
            final RandomAccessible<ExcelCell> ra = headers.get();
            final ExcelCell[] cells = new ExcelCell[ra.size()];
            for (int i = 0; i < ra.size(); i++) {
                final ExcelCell excelCell = ra.get(i);
                if (excelCell == null || excelCell.getStringValue().trim().isEmpty()) {
                    cells[i] = null;
                } else {
                    cells[i] = excelCell;
                }
            }
            return Optional.of(RandomAccessibleUtils.createFromArrayUnsafe(cells));
        }
        return Optional.of(RandomAccessibleUtils.createFromArrayUnsafe());
    }

    @Override
    public Optional<Path> getPath() {
        return m_read.getPath();
    }

}
