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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
package org.knime.ext.poi.node.write2;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;

import org.knime.core.data.DataTable;
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
import org.knime.core.node.util.CheckUtils;
import org.knime.core.node.util.filter.NameFilterConfiguration.FilterResult;
import org.knime.core.node.util.filter.column.DataColumnSpecFilterConfiguration;
import org.knime.core.util.FileUtil;
import org.knime.ext.poi.POIActivator;

/**
 *
 * @author ohl, University of Konstanz
 */
public class XLSWriter2NodeModel extends NodeModel {

    private XLSNodeType m_type;

    private XLSWriter2Settings m_settings = null;

    private DataColumnSpecFilterConfiguration m_filterConfig = null;

    /**
     * @param type Of what type is this node?
     *
     */
    public XLSWriter2NodeModel(final XLSNodeType type) {
        super(1, 0); // one input, no output
        POIActivator.mkTmpDirRW_Bug3301();
        m_type = type;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final DataTableSpec[] inSpecs) throws InvalidSettingsException {
        if (m_settings == null) {
            m_settings = new XLSWriter2Settings();
        }

        String warning =
            CheckUtils.checkDestinationFile(m_settings.getFilename(), (m_type == XLSNodeType.APPENDER)
                || m_settings.getOverwriteOK(), m_settings.getFileMustExist());
        if (warning != null) {
            setWarningMessage(warning);
        }

        return new DataTableSpec[]{};
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData, final ExecutionContext exec)
            throws Exception {
        if (m_filterConfig == null) {
            m_filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        }
        CheckUtils.checkDestinationFile(m_settings.getFilename(), (m_type == XLSNodeType.APPENDER)
            || m_settings.getOverwriteOK(),
            m_settings.getFileMustExist());

        URL url = FileUtil.toURL(m_settings.getFilename());
        Path localPath = FileUtil.resolveToPath(url);

        if (localPath != null) {
            KeyLocker.lock(localPath);
        }

        try {
            if ((localPath != null) && (m_type == XLSNodeType.WRITER) && m_settings.getOverwriteOK()) {
                Files.deleteIfExists(localPath);
            }

            DataTable table = inData[0];
            ColumnRearranger rearranger = new ColumnRearranger(inData[0].getDataTableSpec());
            FilterResult filter = m_filterConfig.applyTo(inData[0].getDataTableSpec());
            rearranger.keepOnly(filter.getIncludes());
            table = exec.createColumnRearrangeTable(inData[0], rearranger, exec);

            XLSWriter2 xlsWriter = new XLSWriter2(url, m_settings);
            xlsWriter.write(table, exec);
        } finally {
            if (localPath != null) {
                KeyLocker.unlock(localPath);
            }
        }

        return new BufferedDataTable[]{};
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settings = new XLSWriter2Settings(settings);
        DataColumnSpecFilterConfiguration filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        filterConfig.loadConfigurationInModel(settings);
        m_filterConfig = filterConfig;
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
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
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
    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        String filename = new XLSWriter2Settings(settings).getFilename();
        if ((filename == null) || filename.isEmpty()) {
            throw new InvalidSettingsException("No output" + " filename specified.");
        }
        DataColumnSpecFilterConfiguration filterConfig = XLSWriter2NodeDialogPane.createColFilterConf();
        filterConfig.loadConfigurationInModel(settings);
    }

}
