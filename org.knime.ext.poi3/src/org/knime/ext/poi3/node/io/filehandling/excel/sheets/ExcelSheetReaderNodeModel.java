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
 *   Nov 16, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.sheets;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.EnumSet;
import java.util.List;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell.StringCellFactory;
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
import org.knime.core.node.streamable.BufferedDataTableRowOutput;
import org.knime.core.node.streamable.InputPortRole;
import org.knime.core.node.streamable.OutputPortRole;
import org.knime.core.node.streamable.PartitionInfo;
import org.knime.core.node.streamable.PortInput;
import org.knime.core.node.streamable.PortOutput;
import org.knime.core.node.streamable.RowOutput;
import org.knime.core.node.streamable.StreamableOperator;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.data.location.FSLocationValueMetaData;
import org.knime.filehandling.core.data.location.cell.FSLocationCellFactory;
import org.knime.filehandling.core.data.location.cell.MultiFSLocationCellFactory;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.ReadPathAccessor;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;

/**
 * Node model to extract the sheet names from a given excel file.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelSheetReaderNodeModel extends NodeModel {

    private final NodeModelStatusConsumer m_statusConsumer;

    private final SettingsModelReaderFileChooser m_filechooser;

    ExcelSheetReaderNodeModel(final PortsConfiguration portsConfig, final SettingsModelReaderFileChooser fileChooser) {
        super(portsConfig.getInputPorts(), portsConfig.getOutputPorts());
        m_statusConsumer = new NodeModelStatusConsumer(EnumSet.of(MessageType.ERROR, MessageType.WARNING));
        m_filechooser = fileChooser;
    }

    @Override
    protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        m_filechooser.configureInModel(inSpecs, m_statusConsumer);
        m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
        return new DataTableSpec[]{createSpec()};
    }

    private DataTableSpec createSpec() {
        final DataColumnSpecCreator colCreator = new DataColumnSpecCreator("Path", FSLocationCellFactory.TYPE);
        final FSLocation location = m_filechooser.getLocation();
        final FSLocationValueMetaData metaData = new FSLocationValueMetaData(location.getFileSystemCategory(),
            location.getFileSystemSpecifier().orElse(null));
        colCreator.addMetaData(metaData, true);
        return new DataTableSpec(colCreator.createSpec(),
            new DataColumnSpecCreator("Sheet", StringCellFactory.TYPE).createSpec());
    }

    @Override
    protected PortObject[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        final BufferedDataTableRowOutput container =
            new BufferedDataTableRowOutput(exec.createDataContainer(createSpec()));
        try {
            write(exec, container);
        } finally {
            container.close();
        }
        return new BufferedDataTable[]{container.getDataTable()};
    }

    private void write(final ExecutionContext exec, final RowOutput output)
        throws IOException, InvalidSettingsException, InterruptedException, CanceledExecutionException {
        try (ReadPathAccessor accessor = m_filechooser.createReadPathAccessor();
                final MultiFSLocationCellFactory fac = new MultiFSLocationCellFactory()) {
            final List<FSPath> inputPaths = accessor.getFSPaths(m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
            int rowOffset = 0;
            double progressSteps = 1d / inputPaths.size();
            int iter = 1;
            for (final FSPath input : inputPaths) {
                exec.checkCanceled();
                exec.setMessage(() -> String.format("Processing file '%s'", input.toString()));
                rowOffset += writeSheetNames(exec, output, fac, input, rowOffset);
                exec.setProgress(iter * progressSteps);
                ++iter;
            }
        }
    }

    private static int writeSheetNames(final ExecutionContext exec, final RowOutput output,
        final MultiFSLocationCellFactory fac, final FSPath input, final int rowOffset)
        throws InterruptedException, CanceledExecutionException, IOException {
        try (final InputStream in = Files.newInputStream(input);
                final BufferedInputStream buffered = new BufferedInputStream(in)) {
            final FSLocation fsLocation = input.toFSLocation();
            try (final Workbook wb = WorkbookFactory.create(buffered)) {
                return pushRows(exec, fac, fsLocation, output, wb, rowOffset);
            }
        }
    }

    private static int pushRows(final ExecutionContext exec, final MultiFSLocationCellFactory fac,
        final FSLocation fsLocation, final RowOutput output, final Workbook wb, final int rowOffset)
        throws InterruptedException, CanceledExecutionException {
        final int numOfRows = wb.getNumberOfSheets();
        for (int i = 0; i < numOfRows; i++) {
            exec.checkCanceled();
            output.push(new DefaultRow("Row" + (i + rowOffset), fac.createCell(exec, fsLocation),
                StringCellFactory.create(wb.getSheetName(i))));
        }
        return numOfRows;
    }

    @Override
    public InputPortRole[] getInputPortRoles() {
        return Stream.generate(() -> InputPortRole.NONDISTRIBUTED_NONSTREAMABLE)//
            .limit(getNrInPorts())//
            .toArray(InputPortRole[]::new);
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
                final RowOutput output = (RowOutput)outputs[0];
                try {
                    write(exec, output);
                } finally {
                    output.close();
                }
            }
        };
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        m_filechooser.saveSettingsTo(settings);
    }

    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_filechooser.validateSettings(settings);
    }

    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_filechooser.loadSettingsFrom(settings);
    }

    @Override
    protected void reset() {
        // nothing to do
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
}
