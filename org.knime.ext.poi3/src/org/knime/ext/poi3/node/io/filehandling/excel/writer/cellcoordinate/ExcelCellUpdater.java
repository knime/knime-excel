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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.cellcoordinate;

import java.io.IOException;

import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.streamable.RowInput;
import org.knime.ext.poi3.node.io.filehandling.excel.updater.cell.ExcelCellUpdaterConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.updater.cell.ExcelCellUpdaterNodeModel.AppendWorkbookHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.WorkbookHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelProgressMonitor;
import org.knime.filehandling.core.connections.FSPath;

/**
 * This writer updates individual cells in sheets of an excel file and finally stores this excel file to disc.
 *
 * @author Moditha Hewasinghage,, KNIME GmbH, Berlin, Germany
 * @author Jannik Löscher, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelCellUpdater {

    private final ExcelCellUpdaterConfig m_cfg;

    /**
     * Constructor.
     *
     * @param cfg the {@link ExcelTableConfig}
     */
    public ExcelCellUpdater(final ExcelCellUpdaterConfig cfg) {
        m_cfg = cfg;
    }

    /**
     * Writes the {@link RowInput}s to individual cells in sheets of an excel file and finally stores this excel file to
     * disc
     *
     * @param outPath the location the excel file has to be written to
     * @param tables the tables to be written to individual sheets
     * @param coordinateColumnIndices the indices of the columns in the tables containing the coordinates
     * @param wbHandler the {@link WorkbookHandler}
     * @param exec the {@link ExecutionContext}
     * @param m the {@link ExcelProgressMonitor}
     * @throws IOException - If the file could not be written to the output path
     * @throws InvalidSettingsException - If the sheet names cannot be made unique
     * @throws CanceledExecutionException - If the execution was canceled by the user
     * @throws InterruptedException - If the execution was canceled by the user
     */
    public void writeTables(final FSPath outPath, final RowInput[] tables, final int[] coordinateColumnIndices,
        final AppendWorkbookHandler wbHandler, final ExecutionContext exec, final ExcelProgressMonitor m)
        throws IOException, InvalidSettingsException, CanceledExecutionException, InterruptedException {
        @SuppressWarnings("resource") // try-with-resources does not work in case of SXSSFWorkbooks
        final var wb = wbHandler.getWorkbook();
        final var creationHelper = wb.getCreationHelper();
        final var cellWriterFactory =
            ExcelCellWriterFactory.createFactory(wb, m_cfg.getMissingValPattern().orElse(null));

        final String[] sheetNames = m_cfg.getSheetNames();
        for (var i = 0; i < tables.length; i++) {
            exec.checkCanceled();
            final var rowInput = tables[i];
            final var writer = wbHandler.createTableWriter(m_cfg, cellWriterFactory);
            writer.writeCellsFromCoordinates(wb, sheetNames[i], rowInput, coordinateColumnIndices[i], m);
        }
        if (m_cfg.evaluate()) {
            final var formulaCtx = exec.createSubExecutionContext(0.05);
            formulaCtx.setMessage("Evaluating formulas");
            creationHelper.createFormulaEvaluator().evaluateAll();
            formulaCtx.setProgress(1);
        }

        exec.setMessage(String.format("Saving excel file to '%s'", outPath.toString()));
        wbHandler.saveFile(outPath);
        exec.setProgress(1);
    }

}
