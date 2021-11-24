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
 *   Aug 9, 2021 (Moditha Hewasinghage, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.EnumSet;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.StringValue;
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
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cellcoordinate.ExcelCellUpdater;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.WorkbookCreator;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelProgressMonitor;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;

/**
 * {@link NodeModel} writing values to individual cells of an excel file.
 *
 * @author Moditha Hewasinghage, KNIME GmbH, Berlin, Germany
 * @author Jannik Löscher, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelCellUpdaterNodeModel extends NodeModel {

    /** The maximum progress for creating the excel file. */
    private static final double MAX_EXCEL_PROGRESS = 0.75;

    private final ExcelCellUpdaterConfig m_cfg;

    private final int[] m_dataPortIndices;

    private int[] m_coordinateColumnIndices;

    private final NodeModelStatusConsumer m_statusConsumer;

    /**
     * Constructor.
     */
    ExcelCellUpdaterNodeModel(final PortsConfiguration portsConfig) {
        super(portsConfig.getInputPorts(), portsConfig.getOutputPorts());
        m_cfg = new ExcelCellUpdaterConfig(portsConfig);
        m_dataPortIndices = portsConfig.getInputPortLocation().get(ExcelCellUpdaterNodeFactory.SHEET_GRP_ID);
        m_statusConsumer = new NodeModelStatusConsumer(EnumSet.of(MessageType.ERROR, MessageType.WARNING));
        m_coordinateColumnIndices = new int[m_dataPortIndices.length];
    }

    @Override
    protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        m_cfg.getSrcFileChooserModel().configureInModel(inSpecs, m_statusConsumer);
        m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);

        try {
            m_cfg.getDestFileChooserModel().configureInModel(inSpecs, m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
        } catch (InvalidSettingsException e) {
            if (m_cfg.isCreateNewFile().getBooleanValue()) {
                throw e;
            }
        }

        CheckUtils.checkSetting(Arrays.stream(m_cfg.getSheetNames()).allMatch(Objects::nonNull),
            "File could not be read or at least one input table has no sheet name set. Please re-configure the node.");
        checkAllHaveCoordinateColumn(inSpecs);

        return new DataTableSpec[0];
    }

    void checkAllHaveCoordinateColumn(final PortObjectSpec[] specs) throws InvalidSettingsException {
        final var names = m_cfg.getCoordinateColumnNames();
        // this stream has side effects because the #checkExistsOrSelectFirst(…) method updates the names array
        final var illlegalSpecs = IntStream.range(0, m_dataPortIndices.length)//
            .filter(i -> //
            !((DataTableSpec)specs[m_dataPortIndices[i]]).containsCompatibleType(StringValue.class) || //
                !checkExistsOrSaveFirst((DataTableSpec)specs[m_dataPortIndices[i]], names, i))//
            .map(i -> m_dataPortIndices[i])//
            .mapToObj(Integer::toString)//
            .collect(Collectors.joining(", "));
        m_cfg.setCoordinateColumnNames(names);

        CheckUtils.checkSetting(illlegalSpecs.isEmpty(),
            "The tables at the following ports do not have a string column of the configured name for the addresses: %s",
            illlegalSpecs);
        m_coordinateColumnIndices = IntStream.range(0, m_dataPortIndices.length)//
            .map(i -> ((DataTableSpec)specs[m_dataPortIndices[i]]).findColumnIndex(names[i]))//
            .toArray();
    }

    static boolean checkExistsOrSaveFirst(final DataTableSpec spec, final String[] names, final int index) {
        if (spec.containsName(names[index])) { // NOSONAR: False positive: would unexpectedly update array
            return true;
        }

        final var noSelection = names[index] == null || names[index].isEmpty();
        for (final var cspec : spec) {
            if (cspec.getType().isCompatible(StringValue.class)) {
                names[index] = cspec.getName();
                break;
            }
        }

        return noSelection;

    }

    @Override
    protected PortObject[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        final var bufferedTables = Arrays.stream(m_dataPortIndices)//
            .mapToObj(i -> (BufferedDataTable)inObjects[i])//
            .toArray(BufferedDataTable[]::new);
        final long[] tableSizes = Arrays.stream(bufferedTables)//
            .mapToLong(BufferedDataTable::size)//
            .toArray();

        final var excelProgress = new ExcelProgressMonitor(getExcelWriteSubProgress(exec), //
            Arrays.stream(tableSizes).sum());
        final RowInput[] tables = Arrays.stream(bufferedTables)//
            .map(DataTableRowInput::new)//
            .toArray(RowInput[]::new);
        write(exec, excelProgress, tables, m_coordinateColumnIndices);
        return new BufferedDataTable[0];
    }

    private void write(final ExecutionContext exec, final ExcelProgressMonitor m, final RowInput[] tables,
        final int[] coordinateColumnIndices)
        throws InvalidSettingsException, IOException, CanceledExecutionException, InterruptedException {

        final var destFileChooser = m_cfg.getDestFileChooserModel();
        final var srcFileChooser = m_cfg.getSrcFileChooserModel();
        if (!m_cfg.isCreateNewFile().getBooleanValue()) {
            final var dest = srcFileChooser.getLocation();
            destFileChooser.setLocation(dest);
        }

        try (final var destAccessor = destFileChooser.createWritePathAccessor();
                final var readAccessor = srcFileChooser.createReadPathAccessor()) {

            final var inputPath = readAccessor.getRootPath(m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
            final var outputPath = destAccessor.getOutputPath(m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);

            createOutputFoldersIfMissing(outputPath.toAbsolutePath().getParent(),
                destFileChooser.isCreateMissingFolders());
            checkOutputFile(outputPath, destFileChooser.getFileOverwritePolicy());

            exec.setMessage("Opening excel file");
            final var wbCreator = getWorkbookCreator(inputPath);
            final var writer = new ExcelCellUpdater(m_cfg);
            writer.writeTables(outputPath, tables, coordinateColumnIndices, wbCreator, exec, m);
        }
    }

    private static ExecutionContext getExcelWriteSubProgress(final ExecutionContext exec) {
        return exec.createSubExecutionContext(MAX_EXCEL_PROGRESS);
    }

    private static AppendWorkbookCreator getWorkbookCreator(final Path path) {
        return new AppendWorkbookCreator(path);
    }

    private static void createOutputFoldersIfMissing(final Path outputFolder, final boolean createMissingFolders)
        throws IOException {
        if (!FSFiles.exists(outputFolder)) {
            if (createMissingFolders) {
                FSFiles.createDirectories(outputFolder);
            } else {
                throw new IOException(String.format(
                    "The directory '%s' does not exist and must not be created due to user settings.", outputFolder));
            }
        }
    }

    private void checkOutputFile(final Path outputFile, final FileOverwritePolicy policy)
        throws InvalidSettingsException, IOException {
        CheckUtils.checkSetting(!m_cfg.isCreateNewFile().getBooleanValue() || policy != FileOverwritePolicy.FAIL
            || !FSFiles.exists(outputFile), "The output file '%s' already exists and must not be overwritten due to user settings.", outputFile);
    }

    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec) {
        // nothing to do
    }

    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec) {
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
        final var inputRoles = new InputPortRole[numInputPorts];
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
                final var excelProgress = new ExcelProgressMonitor(getExcelWriteSubProgress(exec));
                final var tables = Arrays.stream(m_dataPortIndices)//
                    .mapToObj(i -> (RowInput)inputs[i])//
                    .toArray(RowInput[]::new);
                write(exec, excelProgress, tables, m_coordinateColumnIndices);
            }
        };
    }

    static ExcelFormat getExcelFormat(final String path) {
        final var extension = FilenameUtils.getExtension(path);
        switch (extension) {
            case "xls":
                return ExcelFormat.XLS;
            case "xlsx":
                return ExcelFormat.XLSX;
            default:
                return null;
        }
    }

    /**
     * {@link WorkbookCreator} for existing excel files.
     *
     * @author Moditha Hewasinghage, KNIME GmbH, Berlin, Germany
     */
    public static class AppendWorkbookCreator implements WorkbookCreator {

        private final ExcelFormat m_format;

        private final Path m_path;

        AppendWorkbookCreator(final Path path) {
            m_path = path;
            m_format = getExcelFormat(path.getFileName().toString());
        }

        @Override
        public Workbook createWorkbook() throws IOException {
            final var input = new BufferedInputStream(Files.newInputStream(m_path));
            return WorkbookFactory.create(input);
        }

        @Override
        public ExcelTableWriter createTableWriter(final ExcelTableConfig cfg,
            final ExcelCellWriterFactory cellWriterFactory) {
            CheckUtils.checkState(m_format != null, "Cannot create a table writer before creating a workbook");
            return m_format.createWriter(cfg, cellWriterFactory);
        }

    }

}
