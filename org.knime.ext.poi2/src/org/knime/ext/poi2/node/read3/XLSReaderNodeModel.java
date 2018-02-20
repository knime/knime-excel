/*
 * ------------------------------------------------------------------------
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
 * -------------------------------------------------------------------
 *
 * History
 *   Apr 8, 2009 (ohl): created
 */
package org.knime.ext.poi2.node.read3;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Locale;
import java.util.concurrent.ExecutionException;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTable;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeCreationContext;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.streamable.OutputPortRole;
import org.knime.core.node.streamable.PartitionInfo;
import org.knime.core.node.streamable.PortInput;
import org.knime.core.node.streamable.PortOutput;
import org.knime.core.node.streamable.RowOutput;
import org.knime.core.node.streamable.StreamableOperator;
import org.knime.core.node.streamable.StreamableOperatorInternals;
import org.knime.core.util.FileUtil;
import org.xml.sax.SAXException;

/**
 * Xls(x) reader node's {@link NodeModel}.
 *
 * @author Peter Ohl, KNIME AG, Zurich, Switzerland
 * @author Gabor Bakos
 */
public class XLSReaderNodeModel extends NodeModel {

    /**
     * An implementation of {@link StreamableOperator} for streaming Excel reading.
     */
    private final class XLSReaderStreamableOperator extends StreamableOperator {
        private DataTable m_dataTable;
        private CachedExcelTable m_table;

        /**
         * {@inheritDoc}
         */
        @Override
        public void runIntermediate(final PortInput[] inputs, final ExecutionContext exec) throws Exception {
            computeDataTable(new ExecutionMonitor());
            m_dts = m_dataTable.getDataTableSpec();
        }

        @Override
        public void runFinal(final PortInput[] inputs, final PortOutput[] outputs, final ExecutionContext exec)
            throws Exception {
            //With the current implementation of CachedExcelTable it does not make sense to create an intermediate step
            //as it would require serializing and deserializing the whole sheet
            // Changes here should probably also be applied to #execute
            if (m_dataTable == null) {
                computeDataTable(exec);
            }
            final RowOutput output = (RowOutput)outputs[0];
            long curRow = 0L;
            final long totalRow = m_table.lastRow();
            ExecutionContext toTableExec = exec.createSubExecutionContext(.1);
            for (final DataRow row : m_dataTable) {
                final long curRowFinal = curRow;
                toTableExec.setProgress(curRow / (double)totalRow,
                    () -> String.format("Row %d (\"%s\")", curRowFinal, totalRow));
                curRow += 1L;
                toTableExec.checkCanceled();
                output.push(row);
            }
            output.close();
        }

        /**
         * @throws InvalidSettingsException
         * @throws ParserConfigurationException
         * @throws OpenXML4JException
         * @throws SAXException
         * @throws IOException
         * @throws InvalidFormatException
         * @throws ExecutionException
         * @throws InterruptedException
         */
        private void computeDataTable(final ExecutionMonitor exec) throws InvalidSettingsException, InvalidFormatException, IOException, SAXException, OpenXML4JException, ParserConfigurationException, InterruptedException, ExecutionException {
            // Execute is an isolated call so do not keep the workbook in memory
            // Changes here should probably also be applied to #execute
            final XLSUserSettings settings = XLSUserSettings.clone(m_settings);
            final String loc = settings.getFileLocation();
            String sheetName = settings.getSheetName();
            sheetName = settings(settings, sheetName, loc);
            settings.setSheetName(sheetName);
            try (final InputStream is = FileUtil.openInputStream(loc, 1000 * settings.getTimeoutInSeconds())) {
                m_table = isXlsx() && !settings.isReevaluateFormulae()
                    ? CachedExcelTable.fillCacheFromXlsxStreaming(is, sheetName, Locale.ENGLISH,
                        exec, null).get()
                    : CachedExcelTable.fillCacheFromDOM(loc, is, sheetName, Locale.ENGLISH,
                        settings.isReevaluateFormulae(), exec, null).get();
            }
            m_dataTable = m_table.createDataTable(settings, null);
        }
    }

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSReaderNodeModel.class);

    private XLSUserSettings m_settings = new XLSUserSettings();

    private DataTableSpec m_dts = null;

    private String m_dtsSettingsID = null;

    /**
     *
     */
    public XLSReaderNodeModel() {
        super(0, 1);
    }

    XLSReaderNodeModel(final NodeCreationContext context) {
        this();
        m_settings.setFileLocation(context.getUrl().toString());
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        /*
         * This is a special "deal" for the reader: The reader, if previously
         * executed, has data at it's output - even if the file that was read
         * doesn't exist anymore. In order to warn the user that the data cannot
         * be recreated we check here if the file exists and set a warning
         * message if it doesn't.
         */
        String fName = m_settings.getFileLocation();
        if (fName == null || fName.isEmpty()) {
            return;
        }

        try {
            new URL(fName);
            // don't check URLs - don't open a stream.
            return;
        } catch (MalformedURLException mue) {
            // continue on a file
        }
        File location = new File(fName);

        if (!location.canRead() || location.isDirectory()) {
            setWarningMessage("The file '" + location.getAbsolutePath() + "' can't be accessed anymore!");
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settings = XLSUserSettings.load(settings);
        try {
            String settingsID = settings.getString(XLSReaderNodeDialog.XLS_CFG_ID_FOR_TABLESPEC);
            if (!m_settings.getID().equals(settingsID)) {
                throw new InvalidSettingsException("IDs don't match");
            }
            NodeSettingsRO dtsConfig = settings.getNodeSettings(XLSReaderNodeDialog.XLS_CFG_TABLESPEC);
            m_dts = DataTableSpec.load(dtsConfig);
            m_dtsSettingsID = settingsID;
        } catch (InvalidSettingsException ise) {
            LOGGER.debug("No DTS saved in settings");
            // it's optional - if it's not saved we create it later
            m_dts = null;
            m_dtsSettingsID = null;
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void reset() {
        // empty
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // empty
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        if (m_settings != null) {
            m_settings.save(settings);
            if (m_dts != null) {
                settings.addString(XLSReaderNodeDialog.XLS_CFG_ID_FOR_TABLESPEC, m_dtsSettingsID);
                m_dts.save(settings.addConfig(XLSReaderNodeDialog.XLS_CFG_TABLESPEC));
            }
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        XLSUserSettings.load(settings);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final DataTableSpec[] inSpecs) throws InvalidSettingsException {

        if (m_settings == null) {
            throw new InvalidSettingsException("Node not configured.");
        }

        String errMsg = m_settings.getStatus(true);
        if (errMsg != null) {
            throw new InvalidSettingsException(errMsg);
        }
        if (m_settings.isNoPreview()) {
            return null;
        }

        // make sure the DTS still fits the settings
        if (!m_settings.getID().equals(m_dtsSettingsID)) {
            m_dts = null;
        }
        return new DataTableSpec[]{m_dts};
    }

    /**
     * @param e
     * @return
     */
    String message(final Exception e) {
        String execMsg = e.getMessage();
        if (execMsg == null) {
            execMsg = e.getClass().getSimpleName();
        }
        return execMsg;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData, final ExecutionContext exec)
        throws Exception {
        // As it is possible that we do not have a proper spec, this part does not reuse the logic used in streaming.
        // Execute is an isolated call so do not keep the workbook in memory
        XLSUserSettings settings = XLSUserSettings.clone(m_settings);
        String sheetName = settings.getSheetName();
        String loc = settings.getFileLocation();
        sheetName = settings(settings, sheetName, loc);
        settings.setSheetName(sheetName);
        try (InputStream is = FileUtil.openInputStream(loc, 1000 * settings.getTimeoutInSeconds())) {
            ExecutionContext parseExec = exec.createSubExecutionContext(.9);
            ExecutionContext toTableExec = exec.createSubExecutionContext(.1);
            CachedExcelTable table = isXlsx() && !settings.isReevaluateFormulae()
                ? CachedExcelTable.fillCacheFromXlsxStreaming(is, sheetName, Locale.ENGLISH,
                    parseExec, null).get()
                : CachedExcelTable.fillCacheFromDOM(loc, is, sheetName, Locale.ENGLISH,
                    settings.isReevaluateFormulae(), parseExec, null).get();

            DataTable nonBDTable = table.createDataTable(settings, null);
            BufferedDataContainer container = exec.createDataContainer(nonBDTable.getDataTableSpec());
            long curRow = 0L;
            final long totalRow = table.lastRow();
            for (DataRow r : nonBDTable) {
                container.addRowToTable(r);
                toTableExec.checkCanceled();
                final long curRowFinal = curRow;
                toTableExec.setProgress(curRow / (double)totalRow,
                    () -> String.format("Row %d (\"%s\")", curRowFinal, totalRow));
                curRow += 1L;
            }
            container.close();

            return new BufferedDataTable[]{container.getTable()};
        }
    }

    /**
     * @param settings
     * @param sheetName
     * @param loc
     * @return
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws InvalidFormatException
     */
    private String settings(final XLSUserSettings settings, String sheetName, final String loc)
        throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, InvalidFormatException,
        InvalidSettingsException {
        if (sheetName == null || XLSReaderNodeDialog.FIRST_SHEET.equals(sheetName)) {
            if (isXlsx()) {
                try (InputStream stream =
                        POIUtils.openInputStream(loc, settings.getTimeoutInSeconds());
                        final OPCPackage pack = OPCPackage.open(stream)) {
                    sheetName =
                        POIUtils.getFirstSheetNameWithData(new XSSFReader(pack), new ReadOnlySharedStringsTable(pack));
                }
            } else {
                sheetName = POIUtils.getFirstSheetNameWithData(POIUtils.getWorkbook(loc));
            }
            settings.setSheetName(sheetName);
        }
        return sheetName;
    }

    final boolean isXlsx() {
        return isXlsx(m_settings.getFileLocation());
    }

    /**
     * @param fileName The file name/url
     * @return Whether its name ends with {@code .xsls} (case insensitive).
     */
    static final boolean isXlsx(final String fileName) {
        return fileName.toLowerCase().endsWith(".xlsx");
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public StreamableOperator createStreamableOperator(final PartitionInfo partitionInfo,
        final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        return new XLSReaderStreamableOperator();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean iterate(final StreamableOperatorInternals internals) {
        return m_dts == null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PortObjectSpec[] computeFinalOutputSpecs(final StreamableOperatorInternals internals, final PortObjectSpec[] inSpecs)
        throws InvalidSettingsException {
        PortObjectSpec[] normalSpecs = super.computeFinalOutputSpecs(internals, inSpecs);
        if (normalSpecs == null || normalSpecs[0] == null) {
            if (m_dts != null) {
                return new PortObjectSpec[] {m_dts};
            }
            //Probably this should not happen as during iteration m_dts is computed and it is not ditributable,
            //though tester acts as if it were distributable.
            final XLSUserSettings settings = XLSUserSettings.clone(m_settings);
            final String loc = settings.getFileLocation();
            String sheetName = settings.getSheetName();
            try {
                sheetName = settings(settings, sheetName, loc);
            } catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException ex) {
                throw new RuntimeException(ex);
            }
            settings.setSheetName(sheetName);
            CachedExcelTable table;
            try (final InputStream is = FileUtil.openInputStream(loc, 1000 * settings.getTimeoutInSeconds())) {
                table = isXlsx() && !settings.isReevaluateFormulae()
                    ? CachedExcelTable.fillCacheFromXlsxStreaming(is, sheetName, Locale.ENGLISH,
                        new ExecutionMonitor(), null).get()
                    : CachedExcelTable.fillCacheFromDOM(loc, is, sheetName, Locale.ENGLISH,
                        settings.isReevaluateFormulae(), new ExecutionMonitor(), null).get();
            } catch (IOException | InterruptedException | ExecutionException e) {
                throw new RuntimeException(e);
            }
            final DataTable dataTable = table.createDataTable(settings, null);
            return new PortObjectSpec[] { m_dts = dataTable.getDataTableSpec() };
        }
        return normalSpecs;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public OutputPortRole[] getOutputPortRoles() {
        return new OutputPortRole[]{OutputPortRole.NONDISTRIBUTED};
    }
}
