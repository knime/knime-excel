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
import java.awt.Font;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionListener;
import java.util.Arrays;
import java.util.Map;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.SpinnerNumberModel;
import javax.swing.event.ChangeListener;

import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.ViewUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
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
import org.knime.filehandling.core.util.GBCBuilder;
import org.knime.filehandling.core.util.SettingsUtils;

/**
 * Node dialog of Excel table reader.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class ExcelTableReaderNodeDialog extends AbstractTableReaderNodeDialog<ExcelTableReaderConfig, KNIMECellType> {

    private final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> m_config;

    private final ExcelTableReader m_tableReader;

    private final DialogComponentReaderFileChooser m_filePanel;

    private final SettingsModelReaderFileChooser m_settingsModelFilePanel;

    private final JCheckBox m_use15DigitsPrecision = new JCheckBox("Use Excel 15 digits precision");

    private final JRadioButton m_radioButtonFirstSheetWithData = new JRadioButton(SheetSelection.FIRST.getText(), true);

    private final JRadioButton m_radioButtonSheetByName = new JRadioButton(SheetSelection.NAME.getText());

    private final JRadioButton m_radioButtonSheetByIndex = new JRadioButton(SheetSelection.INDEX.getText());

    private final ButtonGroup m_sheetSelectionButtonGroup = new ButtonGroup();

    private final JLabel m_firstSheetWithDataLabel = new JLabel();

    private JComboBox<String> m_sheetNameSelection = new JComboBox<>();

    private JSpinner m_sheetIndexSelection = new JSpinner(new SpinnerNumberModel(1, 1, Integer.MAX_VALUE, 1));

    private JCheckBox m_failOnDifferingSpecs = new JCheckBox("Fail if specs differ");

    private boolean m_updatingSheetSelection = false;

    ExcelTableReaderNodeDialog(final SettingsModelReaderFileChooser settingsModelReaderFileChooser,
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final ExcelTableReader tableReader,
        final MultiTableReadFactory<ExcelTableReaderConfig, KNIMECellType> readFactory,
        final ProductionPathProvider<KNIMECellType> defaultProductionPathProvider) {
        super(readFactory, defaultProductionPathProvider, true);
        m_settingsModelFilePanel = settingsModelReaderFileChooser;
        m_config = config;
        m_tableReader = tableReader;
        final String[] keyChain =
            Stream.concat(Stream.of("settings"), Arrays.stream(m_settingsModelFilePanel.getKeysForFSLocation()))
                .toArray(String[]::new);
        final FlowVariableModel readFvm = createFlowVariableModel(keyChain, FSLocationVariableType.INSTANCE);
        m_filePanel = new DialogComponentReaderFileChooser(m_settingsModelFilePanel, "excel_reader_writer", readFvm,
            FilterMode.FILE, FilterMode.FILES_IN_FOLDERS);

        m_firstSheetWithDataLabel.setFont(new Font(m_firstSheetWithDataLabel.getFont().getName(), Font.ITALIC,
            m_firstSheetWithDataLabel.getFont().getSize()));
        m_sheetSelectionButtonGroup.add(m_radioButtonFirstSheetWithData);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByName);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByIndex);

        addTab("Settings", createGeneralSettingsTab());
        addTab("Transformation", createTransformationTab());
        addTab("Advanced", createAdvancedSettingsTab());
        registerDialogChangeListeners();
        registerPreviewChangeListeners();
    }

    private void registerDialogChangeListeners() {
        m_radioButtonSheetByName.addChangeListener(l -> m_sheetNameSelection
            .setEnabled(m_radioButtonSheetByName.isEnabled() && m_radioButtonSheetByName.isSelected()));
        m_radioButtonSheetByIndex.addChangeListener(l -> m_sheetIndexSelection
            .setEnabled(m_radioButtonSheetByIndex.isEnabled() && m_radioButtonSheetByIndex.isSelected()));

        m_filePanel.getSettingsModel().getFilterModeModel().addChangeListener(l -> toggleFailOnDifferingCheckBox());

    }

    private void configChanged(final boolean updateSheets) {
        if (updateSheets) {
            m_updatingSheetSelection = true;
            m_tableReader.setChangeListener(l -> {
                m_tableReader.setChangeListener(null);
                ViewUtils.invokeLaterInEDT(this::setSheetNameList);
            });
            m_radioButtonFirstSheetWithData.setSelected(true);
            m_radioButtonSheetByName.setEnabled(false);
            m_radioButtonSheetByIndex.setEnabled(false);
            m_firstSheetWithDataLabel.setText(null);
            configChanged();
        } else {
            if (!m_updatingSheetSelection) {
                m_tableReader.setChangeListener(null);
                configChanged();
            }
        }
    }

    private void setSheetNameList() {
        final Map<String, Boolean> sheetNames = m_tableReader.getSheetNames();
        final String selectedSheet = (String)m_sheetNameSelection.getSelectedItem();
        m_sheetNameSelection.setModel(new DefaultComboBoxModel<>(sheetNames.keySet().toArray(new String[0])));
        if (!sheetNames.isEmpty()) {
            if (selectedSheet != null) {
                m_sheetNameSelection.setSelectedItem(selectedSheet);
            } else {
                m_sheetNameSelection.setSelectedIndex(0);
            }
        } else {
            m_sheetNameSelection.setSelectedIndex(-1);
        }
        final String text = String.format("(%s)", ExcelUtils.getFirstSheetWithDataOrFirstIfAllEmpty(sheetNames));
        m_firstSheetWithDataLabel.setText(text);
        // set the preferred height to be correctly aligned with the radio button
        m_firstSheetWithDataLabel.setPreferredSize(new Dimension((int)new JLabel(text).getPreferredSize().getWidth(),
            (int)m_radioButtonFirstSheetWithData.getPreferredSize().getHeight()));
        final int selectedIndex = (int)m_sheetIndexSelection.getValue();
        m_sheetIndexSelection.setModel(
            new SpinnerNumberModel(selectedIndex > sheetNames.size() ? 1 : selectedIndex, 1, sheetNames.size(), 1));

        m_radioButtonSheetByIndex.setEnabled(true);
        m_radioButtonSheetByName.setEnabled(true);
        m_updatingSheetSelection = false;
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

    private JPanel createSheetSelectionPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Sheet selection"));
        final Insets insets = new Insets(5, 5, 0, 0);
        final GBCBuilder gbcBuilder = new GBCBuilder(insets).resetPos().anchorFirstLineStart();

        panel.add(m_radioButtonFirstSheetWithData, gbcBuilder.build());
        panel.add(m_firstSheetWithDataLabel, gbcBuilder.incX().build());
        panel.add(m_radioButtonSheetByName, gbcBuilder.resetX().incY().setWeightX(0).build());
        panel.add(m_sheetNameSelection, gbcBuilder.incX().build());
        panel.add(m_radioButtonSheetByIndex, gbcBuilder.incY().resetX().build());
        final double smallDouble = 1E-10; // spinner should only fill up space of the combo box, rest is for label
        panel.add(m_sheetIndexSelection, gbcBuilder.incX().setWeightX(smallDouble).fillHorizontal().build());
        panel.add(new JLabel(), gbcBuilder.incX().setWeightX(1 - smallDouble).build());

        return panel;
    }

    private final JPanel createGeneralSettingsTab() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createFilePanel(), gbcBuilder.build());
        panel.add(createSheetSelectionPanel(), gbcBuilder.incY().build());
        panel.add(createPreview(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createAdvancedSettingsTab() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createAdvancedReaderOptionsPanel(), gbcBuilder.build());
        panel.add(createSpecFailingOptionsPanel(), gbcBuilder.incY().build());
        panel.add(createPreview(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createAdvancedReaderOptionsPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Reading options"));
        final GBCBuilder gbcBuilder =
            new GBCBuilder(new Insets(5, 5, 5, 5)).resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(m_use15DigitsPrecision, gbcBuilder.build());
        return panel;
    }

    private JPanel createSpecFailingOptionsPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Options for multiple files"));
        final GBCBuilder gbcBuilder =
            new GBCBuilder(new Insets(5, 5, 5, 5)).resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(m_failOnDifferingSpecs, gbcBuilder.build());
        return panel;
    }

    private void toggleFailOnDifferingCheckBox() {
        final boolean enable = m_filePanel.getSettingsModel().getFilterMode() != FilterMode.FILE;
        m_failOnDifferingSpecs.setEnabled(enable);
    }

    private void registerPreviewChangeListeners() {
        m_settingsModelFilePanel.addChangeListener(l -> configChanged(true));

        final ActionListener actionListener = l -> configChanged(false);
        m_failOnDifferingSpecs.addActionListener(actionListener);
        m_radioButtonFirstSheetWithData.addActionListener(actionListener);
        m_radioButtonSheetByName.addActionListener(actionListener);
        m_radioButtonSheetByIndex.addActionListener(actionListener);
        m_sheetNameSelection.addActionListener(actionListener);
        final ChangeListener changeListener = l -> configChanged(false);
        m_sheetIndexSelection.addChangeListener(changeListener);
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        super.saveSettingsTo(settings);
        m_filePanel.saveSettingsTo(SettingsUtils.getOrAdd(settings, SettingsUtils.CFG_SETTINGS_TAB));
        saveTableReadSettings();
        saveExcelReadSettings();
        m_config.setTableSpecConfig(getTableSpecConfig());
        m_config.saveInDialog(settings);
    }

    @Override
    protected void loadSettings(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        // FIXME: loading should be handled by the config (AP-14460 & AP-14462)
        m_filePanel.loadSettingsFrom(SettingsUtils.getOrEmpty(settings, SettingsUtils.CFG_SETTINGS_TAB), specs);
        m_config.loadInDialog(settings, specs);
        loadTableReadSettings();
        loadExcelSettings();
        if (m_config.hasTableSpecConfig()) {
            loadFromTableSpecConfig(m_config.getTableSpecConfig());
        }
    }

    /**
     * Fill in the setting values in {@link TableReadConfig} using values from dialog.
     */
    private void saveTableReadSettings() {
        m_config.setFailOnDifferingSpecs(m_failOnDifferingSpecs.isSelected());
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = m_config.getTableReadConfig();
        tableReadConfig.setUseRowIDIdx(false);
        tableReadConfig.setUseColumnHeaderIdx(true);
        tableReadConfig.setColumnHeaderIdx(0);
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(false);
    }

    /**
     * Fill in the setting values in {@link ExcelTableReaderConfig} using values from dialog.
     */
    private void saveExcelReadSettings() {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(m_use15DigitsPrecision.isSelected());
        if (m_radioButtonSheetByName.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.NAME);
        } else if (m_radioButtonSheetByIndex.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.INDEX);
        } else {
            excelConfig.setSheetSelection(SheetSelection.FIRST);
        }
        excelConfig.setSheetName((String)m_sheetNameSelection.getSelectedItem());
        excelConfig.setSheetIdx((int)m_sheetIndexSelection.getValue());
    }

    /**
     * Fill in dialog components with {@link TableReadConfig} values.
     */
    private void loadTableReadSettings() {
        m_failOnDifferingSpecs.setSelected(m_config.failOnDifferingSpecs());
    }

    /**
     * Fill in dialog components with {@link ExcelTableReaderConfig} values.
     */
    private void loadExcelSettings() {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        m_use15DigitsPrecision.setSelected(excelConfig.isUse15DigitsPrecision());
        switch (excelConfig.getSheetSelection()) {
            case NAME:
                m_radioButtonSheetByName.setSelected(true);
                break;
            case INDEX:
                m_radioButtonSheetByIndex.setSelected(true);
                break;
            case FIRST:
            default:
                m_radioButtonFirstSheetWithData.setSelected(true);
                break;
        }
        m_sheetNameSelection.addItem(excelConfig.getSheetName());
        m_sheetNameSelection.setSelectedItem(excelConfig.getSheetName());
        m_sheetIndexSelection.setValue(excelConfig.getSheetIdx());
    }

    @Override
    protected MultiTableReadConfig<ExcelTableReaderConfig> getConfig() throws InvalidSettingsException {
        saveTableReadSettings();
        saveExcelReadSettings();
        return m_config;
    }

    @Override
    protected ReadPathAccessor createReadPathAccessor() {
        return m_settingsModelFilePanel.createReadPathAccessor();
    }
}
