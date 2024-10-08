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
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.EnumSet;
import java.util.Optional;

import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
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
import org.knime.core.util.DesktopUtil;
import org.knime.core.util.FileUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelMultiRowInputWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.WorkbookHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelProgressMonitor;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetUtils;
import org.knime.filehandling.core.connections.FSCategory;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.connections.uriexport.URIExporter;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.WritePathAccessor;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;
import org.knime.filehandling.core.util.CheckNodeContextUtil;

/**
 * {@link NodeModel} writing tables to individual sheets of an excel file.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterNodeModel extends NodeModel {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelTableWriterNodeModel.class);

    /** The maximum progress for creating the excel file. */
    private static final double MAX_EXCEL_PROGRESS = 0.75;

    private final ExcelTableWriterConfig m_cfg;

    private final int[] m_dataPortIndices;

    private final NodeModelStatusConsumer m_statusConsumer;

    /**
     * Constructor.
     */
    ExcelTableWriterNodeModel(final PortsConfiguration portsConfig) {
        super(portsConfig.getInputPorts(), portsConfig.getOutputPorts());
        m_cfg = new ExcelTableWriterConfig(portsConfig);
        m_dataPortIndices = portsConfig.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID);
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
        final long[] tableSizes = Arrays.stream(bufferedTables)//
            .mapToLong(BufferedDataTable::size)//
            .toArray();
        // validate sheet names
        validateSheetNames(tableSizes);
        final var m = new ExcelProgressMonitor(getExcelWriteSubProgress(exec), Arrays.stream(tableSizes).sum());
        final RowInput[] tables = Arrays.stream(bufferedTables)//
            .map(DataTableRowInput::new)//
            .toArray(RowInput[]::new);
        write(exec, m, tables);
        return new PortObject[]{};
    }

    private void validateSheetNames(final long[] tableSizes) throws InvalidSettingsException {
        final long maxRowsPerSheet = m_cfg.getExcelFormat().getMaxNumRowsPerSheet();
        final String[] sheetNames = m_cfg.getSheetNames();
        for (int i = 0; i < tableSizes.length; i++) {
            final long rowsToWrite = getRowsToWrite(tableSizes[i], maxRowsPerSheet);
            if (rowsToWrite > maxRowsPerSheet) {
                long numAdditionalSheets = rowsToWrite / maxRowsPerSheet;
                SheetUtils.createUniqueSheetName(sheetNames[i], numAdditionalSheets);
            }
        }
    }

    private long getRowsToWrite(final long tableSize, final long maxRowsPerSheet) {
        final int headerOffset = m_cfg.writeColHeaders() ? 1 : 0;
        final long sheetsToWrite = Math.max(1, tableSize / maxRowsPerSheet);
        return tableSize + sheetsToWrite * headerOffset;
    }

    private void write(final ExecutionContext exec, final ExcelProgressMonitor m, final RowInput[] tables)
        throws InvalidSettingsException, IOException, CanceledExecutionException, InterruptedException {
        validateColumnCount(tables);
        final SettingsModelWriterFileChooser fileChooser = m_cfg.getFileChooserModel();
        try (final WritePathAccessor accessor = fileChooser.createWritePathAccessor()) {
            final FSPath outputPath = accessor.getOutputPath(m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
            createOutputFoldersIfMissing(outputPath.toAbsolutePath().getParent(), fileChooser.isCreateMissingFolders());
            exec.setMessage("Opening excel file");
            final var writer = new ExcelMultiRowInputWriter(m_cfg);
            try (final var wbHandler = getWorkbookHandler(fileChooser.getFileOverwritePolicy(), outputPath)) {
                writer.writeTables(outputPath, tables, wbHandler, exec, m);
            }

            if (m_cfg.getOpenFileAfterExecModel().getBooleanValue() && !isHeadlessOrRemote()
                && categoryIsSupported(outputPath.toFSLocation().getFSCategory())) {
                openFile(fileChooser, outputPath);
            }
        }
    }

    private void openFile(final SettingsModelWriterFileChooser fileChooser, final FSPath outputPath)
        throws IOException {
        try (final FSConnection connection = fileChooser.getConnection()) {
            final Optional<File> file = toFile(outputPath, connection);
            if (file.isPresent()) {
                DesktopUtil.open(file.get());
            } else {
                setWarningMessage("Non local files cannot be opened");
            }
        }
    }

    static boolean isHeadlessOrRemote() {
        return Boolean.getBoolean("java.awt.headless") || CheckNodeContextUtil.isRemoteWorkflowContext();
    }

    static boolean categoryIsSupported(final FSCategory fsCategory) {
        return fsCategory == FSCategory.LOCAL //
            || fsCategory == FSCategory.RELATIVE //
            || fsCategory == FSCategory.CUSTOM_URL;
    }

    private static Optional<File> toFile(final FSPath outputPath, final FSConnection fsConnection) {
        final FSCategory fsCategory = outputPath.toFSLocation().getFSCategory();
        if (fsCategory != FSCategory.CUSTOM_URL && fsCategory != FSCategory.RELATIVE) {
            return Optional.of(outputPath.toAbsolutePath().toFile());
        }
        try {
            final URIExporter uriExporter = fsConnection.getDefaultURIExporterFactory().getExporter();
            final String uri = uriExporter.toUri(outputPath).toString();
            final URL url = FileUtil.toURL(uri);
            return Optional.ofNullable(FileUtil.getFileFromURL(url));
        } catch (final MalformedURLException | IllegalArgumentException | URISyntaxException e) {
            LOGGER.debug("Unable to resolve custom URL", e);
            return Optional.empty();
        }
    }

    private void validateColumnCount(final RowInput[] tables) {
        final ExcelFormat excelFormat = m_cfg.getExcelFormat();
        final int maxCols = excelFormat.getMaxNumCols();
        final int rowIdxColOffset = m_cfg.writeRowKey() ? 1 : 0;
        for (int i = 0; i < tables.length; i++) {
            CheckUtils.checkArgument(tables[i].getDataTableSpec().getNumColumns() + rowIdxColOffset <= maxCols,
                "The input table at port %d contains exeeds the column limit (%d) for %s.", m_dataPortIndices[i],
                maxCols, excelFormat.name());
        }
    }

    private static ExecutionContext getExcelWriteSubProgress(final ExecutionContext exec) {
        return exec.createSubExecutionContext(MAX_EXCEL_PROGRESS);
    }

    private WorkbookHandler getWorkbookHandler(final FileOverwritePolicy fileOverwritePolicy, final Path path)
            throws IOException {
        final boolean fileExists = FSFiles.exists(path);
        final var format = m_cfg.getExcelFormat();

        final var auth = m_cfg.getAuthentication();
        final var pw = CryptUtil.getPassword(auth, getCredentialsProvider());
        if (fileExists && fileOverwritePolicy == FileOverwritePolicy.APPEND) {
            return new WriteWorkbookHandler(format, path, m_cfg.evaluate(), pw);
        } else {
            if (fileExists && fileOverwritePolicy == FileOverwritePolicy.FAIL) {
                throw new IOException(String.format(
                    "Output file '%s' exists and must not be overwritten due to user settings.", path.toString()));
            }
            return new WriteWorkbookHandler(format, m_cfg.evaluate(), pw);
        }
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

    /**
     * {@link WorkbookHandler} for existing excel files.
     *
     * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
     */
    // visible for testing
    static class WriteWorkbookHandler extends WorkbookHandler {

        private final boolean m_evaluateFormulas;

        /**
         * Constructor.
         *
         * @param path path to the excel file
         */
        WriteWorkbookHandler(final ExcelFormat format, final boolean evaluateFormulas, final String password) {
            this(format, null, evaluateFormulas, password);
        }

        WriteWorkbookHandler(final ExcelFormat format, final Path path, final boolean evaluateFormulas,
                final String password) {
            super(format, path, password);
            m_evaluateFormulas = evaluateFormulas;
        }

        @Override
        @SuppressWarnings("resource")
        public Workbook createWorkbook(final String secretPassword) throws IOException {
            // Overwrite/create file
            if (m_inputPath == null) {
                return m_format.getWorkbook();
            }

            // Append to existing file
            final var formatIsXLS = m_format == ExcelFormat.XLS;
            final var formatIsXLSX = m_format == ExcelFormat.XLSX;
            if (!(formatIsXLS || formatIsXLSX)) {
                throw new IOException(
                    String.format("Unsupported format: Unable to append spreadsheets to %s", m_inputPath.toString()));
            }
            // if create fails the input stream gets closed otherwise it's closed when invoking close on the workbook
            final var bufferedInputStream = new BufferedInputStream(Files.newInputStream(m_inputPath));
            try {
                final var wb = WorkbookFactory.create(bufferedInputStream, secretPassword);

                // Check that the loaded workbook's format equals the expected format
                if (wb instanceof HSSFWorkbook) {
                    if (formatIsXLSX) {
                        wb.close();
                        throw new IOException(String.format(
                            "Wrong format: The file '%s' is xls instead of xlsx or xlsm formatted. Please adapt the "
                                + "configuration", m_inputPath));
                    }
                } else if (wb instanceof XSSFWorkbook xssfWorkbook) {
                    if (formatIsXLS) {
                        wb.close();
                        throw new IOException(String.format(
                            "Wrong format: The file '%s' is xlsx or xlsm instead of xls formatted. Please adapt the "
                                + "configuration", m_inputPath));
                    }
                    // if we need to evaluate formulas we cannot use the faster streaming model (SXSSF), but have to use
                    // the XSSF model, which keeps everything in memory
                    if (!m_evaluateFormulas) {
                        final var streamingWorkbook = new SXSSFWorkbook(xssfWorkbook);
                        // AP-23021: Zip64Mode must not be Always (POI default since POI 5.x),
                        // since Power BI does not support ZIP64 format. Restore POI 4.x default.
                        streamingWorkbook.setZip64Mode(Zip64Mode.AsNeeded);
                        return streamingWorkbook;
                    }
                } else {
                    wb.close();
                    throw new IOException(String.format("Unsupported format: Unable to append spreadsheets to %s",
                        m_inputPath.toString()));
                }
                return wb;
            } catch (final EncryptedDocumentException e) {
                final var msg = "The file to append to is encrypted, but "
                        + (secretPassword == null ? "no password was supplied." : "supplied password is invalid.");
                throw new EncryptedDocumentException(msg, e);
            }
        }

    }
}
