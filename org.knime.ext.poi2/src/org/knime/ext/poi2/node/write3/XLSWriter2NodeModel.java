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
 *   Mar 15, 2007 (ohl): created
 */
package org.knime.ext.poi2.node.write3;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

import javax.swing.filechooser.FileSystemView;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.SystemUtils;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.container.ColumnRearranger;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.filter.NameFilterConfiguration.FilterResult;
import org.knime.core.node.util.filter.column.DataColumnSpecFilterConfiguration;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.defaultnodesettings.FileChooserHelper;
import org.knime.filehandling.core.defaultnodesettings.SettingsModelFileChooser2;
import org.knime.filehandling.core.port.FileSystemPortObject;

/**
 *
 * @author ohl, University of Konstanz
 */
final class XLSWriter2NodeModel extends NodeModel {

    private static final String CFG_KEY_FILE_CHOOSER = "FILE_CHOOSER";

    static final SettingsModelFileChooser2 createSettingsModelFileChooser(final String[] extensions) {
        return new SettingsModelFileChooser2(CFG_KEY_FILE_CHOOSER, extensions);
    }

    private final SettingsModelFileChooser2 m_settingsModel;

    private final XLSNodeType m_type;

    private XLSWriter2Settings m_settings = null;

    private DataColumnSpecFilterConfiguration m_filterConfig = null;

    /**
     * @param creationConfig
     * @param type Of what type is this node?
     *
     */
    XLSWriter2NodeModel(final NodeCreationConfiguration creationConfig, final XLSNodeType type) {
        super(creationConfig.getPortConfig().get().getInputPorts(),
            creationConfig.getPortConfig().get().getOutputPorts());
        m_type = type;
        m_settingsModel = createSettingsModelFileChooser(
            m_type == XLSNodeType.APPENDER ? new String[]{"xlsx", "xls", "xlsm"} : new String[]{"xlsx", "xls"});
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        if (m_settings == null) {
            m_settings = new XLSWriter2Settings();
        }

        return new DataTableSpec[]{};
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final PortObject[] inData, final ExecutionContext exec)
        throws Exception {
        final Optional<FSConnection> fs = FileSystemPortObject.getFileSystemConnection(inData, 1);

        final FileChooserHelper helper;
        try {
            helper = getFileChooserHelper(fs, m_settingsModel);
        } catch (final IOException e) {
            throw new InvalidSettingsException(e.getMessage());
        }
        final Path path = helper.getPathFromSettings();

        if (m_filterConfig == null) {
            m_filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        }

        checkDestinationFile(path, (m_type == XLSNodeType.APPENDER) || m_settings.getOverwriteOK(),
            m_settings.getFileMustExist());

        KeyLocker.lock(path);
        try {
            if ((m_type == XLSNodeType.WRITER) && m_settings.getOverwriteOK()) {
                Files.deleteIfExists(path);
            }
            final BufferedDataTable dataTable = (BufferedDataTable)inData[0];
            final ColumnRearranger rearranger = new ColumnRearranger(dataTable.getDataTableSpec());
            final FilterResult filter = m_filterConfig.applyTo(dataTable.getDataTableSpec());
            rearranger.keepOnly(filter.getIncludes());
            final BufferedDataTable table = exec.createColumnRearrangeTable(dataTable, rearranger, exec);

            final XLSWriter2 xlsWriter = new XLSWriter2(path, m_settings, m_settingsModel.getFileSystemChoice());
            xlsWriter.write(table, table.getDataTableSpec(), (int)table.size(), exec);
            return new BufferedDataTable[]{};
        } finally {
            //always unlock the path
            KeyLocker.unlock(path);
        }

    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settingsModel.loadSettingsFrom(settings);
        m_settings = new XLSWriter2Settings(settings);
        //AP-6408 ... do we really want to do this????
        String filename = m_settingsModel.getPathOrURL();
        if ((filename != null) && !FilenameUtils.getBaseName(filename).trim().isEmpty()
            && FilenameUtils.getExtension(filename).trim().isEmpty()) {
            filename += ".xlsx";
            m_settingsModel.setPathOrURL(filename);
        }
        final DataColumnSpecFilterConfiguration filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        filterConfig.loadConfigurationInModel(settings);
        m_filterConfig = filterConfig;
    }

    private static FileChooserHelper getFileChooserHelper(final Optional<FSConnection> fs,
        final SettingsModelFileChooser2 fileChooserModel) throws IOException {
        return new FileChooserHelper(fs, fileChooserModel);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void reset() {
        // nothing to reset
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        m_settingsModel.saveSettingsTo(settings);
        if (m_settings != null) {
            m_settings.saveSettingsTo(settings);
        }
        if (m_filterConfig == null) {
            m_filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        }
        m_filterConfig.saveConfiguration(settings);
    }

    /**
     * {@inheritDoc}
     */
    @SuppressWarnings("unused")
    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settingsModel.validateSettings(settings);
        new XLSWriter2Settings(settings);
        final DataColumnSpecFilterConfiguration filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        filterConfig.loadConfigurationInModel(settings);
    }

    private static String checkDestinationFile(final Path path, final boolean allowOverwrite,
        final boolean mustExistForAppend) throws InvalidSettingsException {
        if ((path == null) || path.toString().isEmpty()) {
            throw new InvalidSettingsException("No destination location provided! Please enter a valid location.");
        }

        if (Files.exists(path)) {
            if (Files.isDirectory(path)) {
                throw new InvalidSettingsException("Output location '" + path + "' is a directory");
            } else if (!Files.isWritable(path) && !(looksLikeUNC(path) || isWindowsNetworkMount(path))) {
                throw new InvalidSettingsException("Output file '" + path + "' is not writable");
            } else if (mustExistForAppend) {
                return null; // everything OK
            } else if (allowOverwrite) {
                return "Output file '" + path + "' exists and will be overwritten";
            } else {
                throw new InvalidSettingsException(
                    "Output file '" + path + "' exists and must not be overwritten due to user settings");
            }
        } else if (mustExistForAppend) {
            throw new InvalidSettingsException("Output location '" + path + "' does not exist, append not possible");
        } else {
            final Path parent = path.getParent();
            if (parent != null) {
                if (!Files.exists(parent)) {
                    throw new InvalidSettingsException("Directory '" + parent + "' of output file does not exist");
                } else if (!Files.isWritable(path.getParent())
                    && !(looksLikeUNC(path.getParent()) || isWindowsNetworkMount(path))) {
                    throw new InvalidSettingsException("Directory '" + parent + "' is not writable");
                }
            }
        }

        return null;
    }

    private static boolean looksLikeUNC(final Path path) {
        // Java does not handle UNC URLs correctly, see bug #5864
        return SystemUtils.IS_OS_WINDOWS && path.toString().startsWith("\\\\");
    }

    private static boolean isWindowsNetworkMount(final Path path) {
        if (!SystemUtils.IS_OS_WINDOWS) {
            return false;
        }
        final Path root = path.getRoot();
        final FileSystemView fsv = FileSystemView.getFileSystemView();
        return fsv.getSystemDisplayName(root.toFile()).contains("\\\\");
    }

}
