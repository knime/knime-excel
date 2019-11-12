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
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

import org.knime.base.node.io.filehandling.FileHandlingUtil;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.streamable.OutputPortRole;
import org.knime.core.node.streamable.PartitionInfo;
import org.knime.core.node.streamable.PortInput;
import org.knime.core.node.streamable.PortOutput;
import org.knime.core.node.streamable.RowOutput;
import org.knime.core.node.streamable.StreamableOperator;
import org.knime.core.node.streamable.StreamableOperatorInternals;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.defaultnodesettings.FileChooserHelper;
import org.knime.filehandling.core.defaultnodesettings.SettingsModelFileChooser2;
import org.knime.filehandling.core.port.FileSystemPortObjectSpec;

/**
 * Xls(x) reader node's {@link NodeModel}.
 *
 * @author Peter Ohl, KNIME AG, Zurich, Switzerland
 * @author Gabor Bakos
 */
class XLSReaderNodeModel extends NodeModel {

    /**
     * An implementation of {@link StreamableOperator} for streaming Excel reading.
     */
    private final class XLSReaderStreamableOperator extends StreamableOperator {
        private FileHandlingUtil m_util;

        /**
         * {@inheritDoc}
         */
        @Override
        public void runIntermediate(final PortInput[] inputs, final ExecutionContext exec) throws Exception {
            m_util = createFileHandlingUtil();
            m_dts = m_util.createDataTableSpec();
        }

        @Override
        public void runFinal(final PortInput[] inputs, final PortOutput[] outputs, final ExecutionContext exec)
            throws Exception {
            //With the current implementation of CachedExcelTable it does not make sense to create an intermediate step
            //as it would require serializing and deserializing the whole sheet
            // Changes here should probably also be applied to #execute
            if (m_util == null) {
                m_util = createFileHandlingUtil();
            }
            final RowOutput output = (RowOutput)outputs[0];
            m_util.pushRowsToOutput(output, exec);
            output.close();
        }
    }

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSReaderNodeModel.class);

    private XLSUserSettings m_settings = new XLSUserSettings();

    private static final String CFG_KEY_FILE_CHOOSER = "FILE_CHOOSER";

    static final SettingsModelFileChooser2 getSettingsModelFileChooser() {
        return new SettingsModelFileChooser2(CFG_KEY_FILE_CHOOSER, XLSUserSettings.CFG_XLS_LOCATION , "xls" ,"xlsx");
    }

    private final SettingsModelFileChooser2 m_settingsModelFileChooser = getSettingsModelFileChooser();

    private DataTableSpec m_dts = null;

    private String m_dtsSettingsID = null;

    private Optional<FSConnection> m_fs;

    XLSReaderNodeModel(final NodeCreationConfiguration creationConfig) {
        super(creationConfig.getPortConfig().get().getInputPorts(),
            creationConfig.getPortConfig().get().getOutputPorts());
        if (creationConfig.getURLConfig().isPresent()) {
            m_settings.setFileLocation(creationConfig.getURLConfig().get().getUrl().toString());
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        final FileHandlingUtil fhUtil = createFileHandlingUtil();
        return new BufferedDataTable[]{fhUtil.createDataTable(exec)};
    }

    private FileHandlingUtil createFileHandlingUtil() throws InvalidSettingsException, IOException {
        final ExcelTableReader reader = new ExcelTableReader(m_settings);

        return new FileHandlingUtil(reader, getFileChooserHelper());
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
        final String fName = m_settings.getFileLocation();
        if ((fName == null) || fName.isEmpty()) {
            return;
        }

        final Path path = getFileChooserHelper().getPathFromSettings();
        if (!Files.exists(path)) {
            setWarningMessage("The file/directory '" + path.toString() + "' can't be accessed anymore!");
        }
    }

    private FileChooserHelper getFileChooserHelper() throws IOException {
        return new FileChooserHelper(m_fs, m_settingsModelFileChooser, m_settings.getTimeoutInSeconds() * 1000);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settings = XLSUserSettings.load(settings);
        m_settingsModelFileChooser.loadSettingsFrom(settings);

        try {
            final String settingsID = settings.getString(XLSReaderNodeDialog.XLS_CFG_ID_FOR_TABLESPEC);
            if (!m_settings.getID().equals(settingsID)) {
                throw new InvalidSettingsException("IDs don't match");
            }
            final NodeSettingsRO dtsConfig = settings.getNodeSettings(XLSReaderNodeDialog.XLS_CFG_TABLESPEC);
            m_dts = DataTableSpec.load(dtsConfig);
            m_dtsSettingsID = settingsID;
        } catch (final InvalidSettingsException ise) {
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
            m_settingsModelFileChooser.saveSettingsTo(settings);
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
        m_settingsModelFileChooser.validateSettings(settings);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final PortObjectSpec[] inSpecs) throws InvalidSettingsException {

        m_fs = FileSystemPortObjectSpec.getFileSystemConnection(inSpecs, 0);

        if ((m_settings == null) || (m_settingsModelFileChooser == null)) {
            throw new InvalidSettingsException("Node not configured.");
        }

        final String errMsg = m_settings.getStatus();
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
    public PortObjectSpec[] computeFinalOutputSpecs(final StreamableOperatorInternals internals,
        final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        final PortObjectSpec[] normalSpecs = super.computeFinalOutputSpecs(internals, inSpecs);
        if ((normalSpecs == null) || (normalSpecs[0] == null)) {
            if (m_dts != null) {
                return new PortObjectSpec[]{m_dts};
            }
            //Probably this should not happen as during iteration m_dts is computed and it is not distributable,
            //though tester acts as if it were distributable.
            try {
                return new PortObjectSpec[]{m_dts = createFileHandlingUtil().createDataTableSpec()};
            } catch (final Exception e) {
                throw new RuntimeException(e);
            }
        }
        return normalSpecs;
    }

    @Override
    public OutputPortRole[] getOutputPortRoles() {
        return new OutputPortRole[]{OutputPortRole.NONDISTRIBUTED};
    }
}
