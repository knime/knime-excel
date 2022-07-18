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

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Collection;
import java.util.EnumSet;
import java.util.LinkedHashSet;
import java.util.stream.Stream;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.eventusermodel.AbortableHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
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
import org.knime.core.util.FileUtil;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.connections.uriexport.URIExporterIDs;
import org.knime.filehandling.core.connections.uriexport.noconfig.NoConfigURIExporterFactory;
import org.knime.filehandling.core.data.location.FSLocationValueMetaData;
import org.knime.filehandling.core.data.location.cell.MultiSimpleFSLocationCellFactory;
import org.knime.filehandling.core.data.location.cell.SimpleFSLocationCellFactory;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;
import org.knime.filehandling.core.util.IOESupplier;

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
        final var colCreator = new DataColumnSpecCreator("Path", SimpleFSLocationCellFactory.TYPE);
        final var location = m_filechooser.getLocation();
        final var metaData = new FSLocationValueMetaData(location.getFileSystemCategory(),
            location.getFileSystemSpecifier().orElse(null));
        colCreator.addMetaData(metaData, true);
        return new DataTableSpec(colCreator.createSpec(),
            new DataColumnSpecCreator("Sheet", StringCellFactory.TYPE).createSpec());
    }

    @Override
    protected PortObject[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        final var container = new BufferedDataTableRowOutput(exec.createDataContainer(createSpec()));
        try {
            write(exec, container);
        } finally {
            container.close();
        }
        return new BufferedDataTable[]{container.getDataTable()};
    }

    private void write(final ExecutionContext exec, final RowOutput output) throws IOException,
        InvalidSettingsException, InterruptedException, CanceledExecutionException {
        try (final var accessor = m_filechooser.createReadPathAccessor();
                final var conn = m_filechooser.getConnection()) {

            final var fac = new MultiSimpleFSLocationCellFactory();
            final var inputPaths = accessor.getFSPaths(m_statusConsumer);
            m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);

            var rowOffset = 0;
            double progressSteps = 1d / inputPaths.size();
            var iter = 1;

            for (final var path : inputPaths) {
                exec.checkCanceled();

                // due to limitations of the POI streaming API we need a local java.io.File
                // to read sheet names in a memory-efficient streaming way. getLocalFile()
                // will fetch a local temp copy of any remote files and the closeable file container
                // makes sure that it gets deleted
                exec.setMessage(() -> String.format("Fetching file '%s'", path.toString()));
                try (var fileContainer = getLocalFile(path, conn)) {
                    exec.setMessage(() -> String.format("Processing file '%s'", path.toString()));
                    final var localFile = fileContainer.retrieveLocalFile();
                    rowOffset += writeSheetNames(exec, output, fac, localFile, path, rowOffset);
                }

                exec.setProgress(iter * progressSteps);
                ++iter;
            }

        } catch (OpenXML4JException | URISyntaxException e) {
            throw new IOException(e);
        }
    }

    private static class LocalFileContainer implements AutoCloseable {

        private final IOESupplier<File> m_localFileSupplier;
        private final boolean m_deleteOnClose;
        private File m_file;

        LocalFileContainer(final IOESupplier<File> localFile, final boolean deleteOnClose) {
            m_localFileSupplier = localFile;
            m_deleteOnClose = deleteOnClose;
            m_file = null; // null until retrieved
        }

        File retrieveLocalFile() throws IOException {
            if (m_file == null) {
                m_file = m_localFileSupplier.get();
            }
            return m_file;
        }

        @Override
        public void close() {
            if (m_deleteOnClose && m_file != null) {
                FSFiles.deleteSafely(m_file.toPath());
            }
        }
    }

    private static final LocalFileContainer getLocalFile(final FSPath path, final FSConnection connection)
        throws URISyntaxException {

        final var factory = connection.getURIExporterFactory(URIExporterIDs.KNIME_FILE);
        if (factory != null) { // we are local
            final var uri = ((NoConfigURIExporterFactory)factory).getExporter().toUri(path);
            return new LocalFileContainer(() -> Paths.get(uri).toFile(), false);
        } else { // we aren't local or don't have access to a local path so we create a local copy
            return new LocalFileContainer(() -> copyToLocalTempFile(path), true);
        }
    }

    private static File copyToLocalTempFile(final FSPath path) throws IOException {
        var tmpFile = FileUtil.createTempFile("knime-excel-snr", "." + FilenameUtils.getExtension(path.getFileName().toString()));
        try {
            Files.copy(path, tmpFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            return tmpFile;
        } catch (IOException e) {
            // cleanup the now obsolete temp file
            FSFiles.deleteSafely(tmpFile.toPath());
            throw e;
        }
    }


    private static int writeSheetNames(final ExecutionContext exec, final RowOutput output,
        final MultiSimpleFSLocationCellFactory fac, final File file, final FSPath path, final int rowOffset)
        throws InterruptedException, CanceledExecutionException, IOException, OpenXML4JException {
        final Collection<String> names;
        switch (FilenameUtils.getExtension(path.toString()).toUpperCase()) {
            case "XLS":
                try (final var fs = new POIFSFileSystem(file, true)) {
                    names = getSheetNamesHSSF(exec, fs);
                }
                break;
            case "XLSB":
                try (final var pak = OPCPackage.open(file, PackageAccess.READ)) {
                    names = getSheetNamesXSSF(exec, new XSSFBReader(pak));
                }
                break;
            default: // "XLSX", "XLSM"
                try (final var pak = OPCPackage.open(file, PackageAccess.READ)) {
                    names = getSheetNamesXSSF(exec, new XSSFReader(pak));
                }
        }
        return pushRows(exec, fac, path.toFSLocation(), output, names, rowOffset);
    }

    private static Collection<String> getSheetNamesHSSF(final ExecutionContext exec, final POIFSFileSystem file)
        throws CanceledExecutionException, IOException {
        final var names = new LinkedHashSet<String>();
        final var req = new HSSFRequest();
        final var events = new HSSFEventFactory();
        final var listener = new AbortableHSSFListener() {
            @Override
            public short abortableProcessRecord(final Record r) throws HSSFUserException {
                try {
                    exec.checkCanceled();
                    names.add(((BoundSheetRecord)r).getSheetname());
                    return 0;
                } catch (final CanceledExecutionException ignored) { // NOSONAR
                    // execution canceled is checked again before the method returns
                    return 1;
                }
            }
        };

        req.addListener(listener, BoundSheetRecord.sid);
        events.processWorkbookEvents(req, file);
        exec.checkCanceled();
        return names;
    }

    private static Collection<String> getSheetNamesXSSF(final ExecutionContext exec, final XSSFReader reader)
        throws CanceledExecutionException, IOException, OpenXML4JException {
        final var names = new LinkedHashSet<String>();
        final var sheets = (XSSFReader.SheetIterator)reader.getSheetsData();
        while (sheets.hasNext()) {
            exec.checkCanceled();
            try (final var dummy = sheets.next()) {
                names.add(sheets.getSheetName());
            }
        }
        return names;
    }

    private static int pushRows(final ExecutionContext exec, final MultiSimpleFSLocationCellFactory fac,
        final FSLocation fsLocation, final RowOutput output, final Collection<String> sheetsNames, final int rowOffset)
        throws CanceledExecutionException, InterruptedException {
        var numOfRows = 0;
        for (final var name : sheetsNames) {
            exec.checkCanceled();
            output.push(new DefaultRow("Row" + (numOfRows + rowOffset), fac.createCell(fsLocation),
                StringCellFactory.create(name)));
            numOfRows++;
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
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec) {
        // nothing to do
    }

    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec) {
        // nothing to do
    }
}
