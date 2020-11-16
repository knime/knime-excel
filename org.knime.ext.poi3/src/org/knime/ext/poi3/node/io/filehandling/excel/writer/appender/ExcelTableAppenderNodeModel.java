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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.appender;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.EnumSet;

import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.streamable.DataTableRowInput;
import org.knime.core.node.streamable.InputPortRole;
import org.knime.core.node.streamable.OutputPortRole;
import org.knime.core.node.streamable.PartitionInfo;
import org.knime.core.node.streamable.PortInput;
import org.knime.core.node.streamable.PortOutput;
import org.knime.core.node.streamable.RowInput;
import org.knime.core.node.streamable.StreamableOperator;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelMultiTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelProgressMonitor;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.ReadPathAccessor;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;

/**
 * {@link NodeModel} writing tables to individual sheets of an existing excel file.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableAppenderNodeModel extends NodeModel {

    /** The maximum progress for creating the excel file. */
    private static final double MAX_EXCEL_PROGRESS = 0.75;

    private final ExcelTableAppenderConfig m_cfg;

    private final int[] m_dataPortIndices;

    private final NodeModelStatusConsumer m_statusConsumer;

    /**
     * Constructor.
     */
    ExcelTableAppenderNodeModel(final PortsConfiguration portsConfig) {
        super(portsConfig.getInputPorts(), portsConfig.getOutputPorts());
        m_cfg = new ExcelTableAppenderConfig(portsConfig);
        m_dataPortIndices = portsConfig.getInputPortLocation().get(ExcelTableAppenderNodeFactory.SHEET_GRP_ID);
        m_statusConsumer = new NodeModelStatusConsumer(EnumSet.of(MessageType.ERROR, MessageType.WARNING));
    }

    @Override
    protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        m_cfg.getFileChooserModel().configureInModel(inSpecs, m_statusConsumer);
        m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
        return new PortObjectSpec[]{};
    }

    @Override
    protected PortObject[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        final BufferedDataTable[] bufferedTables = Arrays.stream(m_dataPortIndices)//
            .mapToObj(i -> (BufferedDataTable)inObjects[i])//
            .toArray(BufferedDataTable[]::new);
        final ExcelProgressMonitor m = new ExcelProgressMonitor(getExcelWriteSubProgress(exec),
            Arrays.stream(bufferedTables)//
                .mapToLong(BufferedDataTable::size)//
                .sum());
        final RowInput[] tables = Arrays.stream(bufferedTables)//
            .map(DataTableRowInput::new)//
            .toArray(RowInput[]::new);
        write(exec, m, tables);
        return new PortObject[]{};
    }

    private void write(final ExecutionContext exec, final ExcelProgressMonitor m, final RowInput[] tables)
        throws InvalidSettingsException, IOException, CanceledExecutionException, InterruptedException {
        final SettingsModelReaderFileChooser fileChooser = m_cfg.getFileChooserModel();
        try (final ReadPathAccessor accessor = fileChooser.createReadPathAccessor()) {
            final FSPath inputPath = accessor.getFSPaths(m_statusConsumer).get(0);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
            final ExcelMultiTableWriter writer =
                new ExcelMultiTableWriter(m_cfg, FileOverwritePolicy.OVERWRITE.getOpenOptions());
            writer.writeTables(inputPath, tables, ExcelTableAppenderConfig.getWorkbookCreator(inputPath), exec, m);
        }
    }

    private static ExecutionContext getExcelWriteSubProgress(final ExecutionContext exec) {
        return exec.createSubExecutionContext(MAX_EXCEL_PROGRESS);
    }

    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // nothing to do
    }

    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // nothing to do
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        m_cfg.saveSettingsForModel(settings);
    }

    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_cfg.validateSettingsForModel(settings);
    }

    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_cfg.loadSettingsForModel(settings);
    }

    @Override
    protected void reset() {
        // nothing to do
    }

    @Override
    public InputPortRole[] getInputPortRoles() {
        final int numInputPorts = getNrInPorts();
        final InputPortRole[] inputRoles = new InputPortRole[numInputPorts];
        Arrays.fill(inputRoles, InputPortRole.NONDISTRIBUTED_STREAMABLE);
        if (numInputPorts > m_dataPortIndices.length) {
            inputRoles[0] = InputPortRole.NONDISTRIBUTED_NONSTREAMABLE;
        }
        return inputRoles;
    }

    @Override
    public OutputPortRole[] getOutputPortRoles() {
        return new OutputPortRole[]{OutputPortRole.DISTRIBUTED};
    }

    @Override
    public StreamableOperator createStreamableOperator(final PartitionInfo partitionInfo,
        final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        return new StreamableOperator() {

            @Override
            public void runFinal(final PortInput[] inputs, final PortOutput[] outputs, final ExecutionContext exec)
                throws Exception {
                final ExcelProgressMonitor m = new ExcelProgressMonitor(getExcelWriteSubProgress(exec));
                final RowInput[] tables = Arrays.stream(m_dataPortIndices)//
                    .mapToObj(i -> (RowInput)inputs[i])//
                    .toArray(RowInput[]::new);
                write(exec, m, tables);
            }

        };
    }
}
