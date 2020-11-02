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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.awt.Dimension;
import java.util.Arrays;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JCheckBox;
import javax.swing.JPanel;

import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.DialogComponentReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.ReadPathAccessor;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.node.table.reader.MultiTableReadFactory;
import org.knime.filehandling.core.node.table.reader.ProductionPathProvider;
import org.knime.filehandling.core.node.table.reader.config.DefaultMultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.MultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.preview.dialog.AbstractTableReaderNodeDialog;
import org.knime.filehandling.core.util.SettingsUtils;

/**
 * Node dialog of Excel table reader.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class ExcelTableReaderNodeDialog extends AbstractTableReaderNodeDialog<ExcelTableReaderConfig, KNIMECellType> {

    private final DialogComponentReaderFileChooser m_filePanel;

    private final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> m_config;

    private final SettingsModelReaderFileChooser m_settingsModelFilePanel;

    private final JCheckBox m_use15DigitsPrecision = new JCheckBox("Use Excel 15 digits precision");

    ExcelTableReaderNodeDialog(final SettingsModelReaderFileChooser settingsModelReaderFileChooser,
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final MultiTableReadFactory<ExcelTableReaderConfig, KNIMECellType> readFactory,
        final ProductionPathProvider<KNIMECellType> defaultProductionPathProvider) {
        super(readFactory, defaultProductionPathProvider, true);
        m_settingsModelFilePanel = settingsModelReaderFileChooser;
        m_config = config;
        final String[] keyChain =
            Stream.concat(Stream.of("settings"), Arrays.stream(m_settingsModelFilePanel.getKeysForFSLocation()))
                .toArray(String[]::new);
        final FlowVariableModel readFvm = createFlowVariableModel(keyChain, FSLocationVariableType.INSTANCE);
        m_filePanel = new DialogComponentReaderFileChooser(m_settingsModelFilePanel, "excel_reader", readFvm,
            FilterMode.FILE, FilterMode.FILES_IN_FOLDERS);

        addTab("Settings", createFilePanel());
        addTab("Advanced", createAdvancedSettingsPanel());
    }

    private JPanel createFilePanel() {
        final JPanel filePanel = new JPanel();
        filePanel.setLayout(new BoxLayout(filePanel, BoxLayout.X_AXIS));
        filePanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Input location"));
        filePanel.setMaximumSize(
            new Dimension(Integer.MAX_VALUE, m_filePanel.getComponentPanel().getPreferredSize().height));
        filePanel.add(m_filePanel.getComponentPanel());
        filePanel.add(Box.createHorizontalGlue());
        return filePanel;
    }

    private JPanel createAdvancedSettingsPanel() {
        final JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_use15DigitsPrecision.getPreferredSize().height));
        panel.add(m_use15DigitsPrecision);
        panel.add(Box.createHorizontalGlue());
        return panel;
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        m_filePanel.saveSettingsTo(SettingsUtils.getOrAdd(settings, SettingsUtils.CFG_SETTINGS_TAB));
        saveTableReadSettings();
        saveExcelReadSettings();
        m_config.saveInDialog(settings);
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        // FIXME: loading should be handled by the config (AP-14460 & AP-14462)
        m_filePanel.loadSettingsFrom(SettingsUtils.getOrEmpty(settings, SettingsUtils.CFG_SETTINGS_TAB), specs);
        m_config.loadInDialog(settings, specs);
        loadTableReadSettings();
        loadExcelSettings();
    }

    /**
     * Fill in the setting values in {@link TableReadConfig} using values from dialog.
     */
    private void saveTableReadSettings() {
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = m_config.getTableReadConfig();
        tableReadConfig.setUseRowIDIdx(false);
        tableReadConfig.setUseColumnHeaderIdx(true);
        tableReadConfig.setColumnHeaderIdx(0);
    }

    /**
     * Fill in the setting values in {@link ExcelTableReaderConfig} using values from dialog.
     */
    private void saveExcelReadSettings() {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(m_use15DigitsPrecision.isSelected());
    }

    /**
     * Fill in dialog components with {@link TableReadConfig} values.
     */
    private static void loadTableReadSettings() {
        // TODO options are added later
    }

    /**
     * Fill in dialog components with {@link ExcelTableReaderConfig} values.
     */
    private void loadExcelSettings() {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        m_use15DigitsPrecision.setSelected(excelConfig.isUse15DigitsPrecision());
    }

    @Override
    protected MultiTableReadConfig<ExcelTableReaderConfig> getConfig() throws InvalidSettingsException {
        saveTableReadSettings();
        return m_config;
    }

    @Override
    protected ReadPathAccessor createReadPathAccessor() {
        return m_settingsModelFilePanel.createReadPathAccessor();
    }
}
