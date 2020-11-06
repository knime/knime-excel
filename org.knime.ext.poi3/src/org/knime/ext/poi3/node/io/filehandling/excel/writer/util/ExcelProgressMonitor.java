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
 *   Nov 9, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.util;

import java.util.function.Supplier;

import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeProgressMonitor;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelMultiTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableWriter;

/**
 * Progress monitor for the {@link ExcelMultiTableWriter}. It allows {@link ExcelTableWriter}s to conveniently update
 * the overall progress.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public final class ExcelProgressMonitor {

    private final ExecutionContext m_exec;

    private final long m_rowCount;

    private double m_curIdx;

    /**
     * Constructor.
     *
     * @param exec the execution context
     */
    public ExcelProgressMonitor(final ExecutionContext exec) {
        this(exec, -1);
    }

    /**
     * Constructor.
     *
     * @param exec the execution context
     * @param rowCount the row count
     */
    public ExcelProgressMonitor(final ExecutionContext exec, final long rowCount) {
        m_exec = exec;
        m_rowCount = rowCount;
        m_curIdx = 0;
    }

    /**
     * Throws an exception in case the users canceled the execution.
     *
     * @see NodeProgressMonitor#checkCanceled()
     * @throws CanceledExecutionException which indicated the execution will be canceled by this call.
     */
    public void checkCanceled() throws CanceledExecutionException {
        m_exec.checkCanceled();
    }

    /**
     * Updates the progress.
     *
     * @param sheetName the name of the sheet currently written to
     * @param rowIdx the index of the row currently written
     */
    public void updateProgress(final String sheetName, final long rowIdx) {
        final Supplier<String> messageSupplier = () -> String.format("Writing sheet '%s' row %d", sheetName, rowIdx);
        if (m_rowCount > 0) {
            ++m_curIdx;
            m_exec.setProgress(m_curIdx / m_rowCount, messageSupplier);
        } else {
            m_exec.setMessage(messageSupplier);
        }
    }

}
