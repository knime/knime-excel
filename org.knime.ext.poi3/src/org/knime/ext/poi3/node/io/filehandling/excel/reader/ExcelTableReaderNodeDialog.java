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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFormattedTextField;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DialogComponentAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.ViewUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.columnnames.ColumnNameMode;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.DialogComponentReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.ReadPathAccessor;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.node.table.reader.MultiTableReadFactory;
import org.knime.filehandling.core.node.table.reader.ProductionPathProvider;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.MultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.preview.dialog.AbstractPathTableReaderNodeDialog;
import org.knime.filehandling.core.node.table.reader.preview.dialog.AnalysisComponentModel;
import org.knime.filehandling.core.node.table.reader.preview.dialog.TableReaderPreviewModel;
import org.knime.filehandling.core.node.table.reader.preview.dialog.TableReaderPreviewView;
import org.knime.filehandling.core.util.GBCBuilder;
import org.knime.filehandling.core.util.SettingsUtils;

/**
 * Node dialog of Excel table reader.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class ExcelTableReaderNodeDialog
    extends AbstractPathTableReaderNodeDialog<ExcelTableReaderConfig, KNIMECellType> {

    private static final String TRANSFORMATION_TAB = "Transformation";

    private final ExcelMultiTableReadConfig m_fileContentPreviewConfig = createFileContentPreviewSettings();

    private final ExcelMultiTableReadConfig m_config;

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

    private JSpinner m_sheetIndexSelection = new JSpinner(new SpinnerNumberModel(0, 0, Integer.MAX_VALUE, 1));

    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_sheetIdxNoteLabel = new JLabel("(Sheet indexes start with 0.) ");

    private JCheckBox m_failOnDifferingSpecs = new JCheckBox("Fail if specs differ");

    private final JCheckBox m_columnHeaderCheckBox = new JCheckBox("Table contains column names in row number", true);

    private final JSpinner m_columnHeaderSpinner = new JSpinner(
        new SpinnerNumberModel(Long.valueOf(1), Long.valueOf(1), Long.valueOf(Long.MAX_VALUE), Long.valueOf(1)));

    private final JLabel m_emptyColHeaderLabel = new JLabel("Empty column name prefix:");

    private final JTextField m_emptyColHeaderPrefix = new JTextField("empty_");

    private final JRadioButton m_emptyColHeaderIndex = new JRadioButton(ColumnNameMode.COL_INDEX.getText());

    private final JRadioButton m_emptyColHeaderExcelColName = new JRadioButton(ColumnNameMode.EXCEL_COL_NAME.getText(), true);

    private final ButtonGroup m_emptyColHeaderBtnGrp = new ButtonGroup();

    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_columnHeaderNoteLabel =
        new JLabel("(Row numbers start with 1. See \"File Content\" tab to identify row numbers.) ");

    private final JCheckBox m_skipHiddenCols = new JCheckBox("Skip hidden columns", true);

    private final JCheckBox m_skipHiddenRows = new JCheckBox("Skip hidden rows", true);

    private final JCheckBox m_skipEmptyRows = new JCheckBox("Skip empty rows", true);

    private final JCheckBox m_skipEmptyCols = new JCheckBox(
        "Skip empty columns (whether a column is considered empty depends on the \"Table specification\" settings "
            + "below)",
        false);

    private final JCheckBox m_replaceEmptyStringsWithMissings =
        new JCheckBox("Replace empty strings with missing values", true);

    private boolean m_updatingSheetSelection = false;

    private final JCheckBox m_reevaluateFormulas =
        new JCheckBox("Reevaluate formulas (leave unchecked if uncertain; see node description for details)");

    private final JRadioButton m_radioButtonInsertErrorPattern =
        new JRadioButton(FormulaErrorHandling.PATTERN.getText(), true);

    private final JRadioButton m_radioButtonInsertMissingCell =
        new JRadioButton(FormulaErrorHandling.MISSING.getText());

    private final ButtonGroup m_buttonGroupFormulaError = new ButtonGroup();

    private final JTextField m_formulaErrorPattern = new JTextField();

    private final JRadioButton m_radioButtonReadEntireSheet =
        new JRadioButton(AreaOfSheetToRead.ENTIRE.getText(), true);

    private final JRadioButton m_radioButtonReadPartOfSheet = new JRadioButton(AreaOfSheetToRead.PARTIAL.getText());

    private final ButtonGroup m_buttonGroupSheetArea = new ButtonGroup();

    private final JLabel m_fromColLabel = new JLabel("columns from");

    private final JLabel m_toColLabel = new JLabel("to");

    private final JLabel m_andLabel = new JLabel("and");

    private final JLabel m_fromRowLabel = new JLabel("rows from");

    private final JLabel m_toRowLabel = new JLabel("to");

    private final JLabel m_dotLabel = new JLabel(".");

    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_readAreaNoteLabel = new JLabel("(See \"File Content\" tab to identify columns and rows.) ");

    private final JTextField m_fromCol = new JTextField("A");

    private final JTextField m_toCol = new JTextField();

    private final JTextField m_fromRow = new JTextField("1");

    private final JTextField m_toRow = new JTextField();

    private final JRadioButton m_radioButtonGenerateRowIDs = new JRadioButton(RowIDGeneration.GENERATE.getText(), true);

    private final JRadioButton m_radioButtonReadRowIDsFromCol = new JRadioButton(RowIDGeneration.COLUMN.getText());

    private final ButtonGroup m_buttonGrouprowIDGeneration = new ButtonGroup();

    private final JTextField m_rowIDColumn = new JTextField("A");

    private final JCheckBox m_limitAnalysisChecker = new JCheckBox("Limit data rows scanned", true);

    private final JSpinner m_limitAnalysisSpinner = new JSpinner(
        new SpinnerNumberModel(Long.valueOf(50), Long.valueOf(1), Long.valueOf(Long.MAX_VALUE), Long.valueOf(50)));

    private final JCheckBox m_supportChangingFileSchemas = new JCheckBox("Support changing file schemas");

    private final DialogComponentAuthentication m_passwordComponent;

    private List<JTabbedPane> m_tabbedPreviewPaneList = new ArrayList<>();

    private final FileContentPreviewController<ExcelTableReaderConfig, KNIMECellType> m_fileContentPreviewController;

    private final List<TableReaderPreviewView> m_fileContentPreviews = new ArrayList<>();

    private final TableReaderPreviewModel m_previewModel;

    private boolean m_fileContentConfigChanged = true;

    private boolean m_previewConfigChanged = false;

    private boolean m_switchTabInTabbedPanes = false;

    private final SettingsModelAuthentication m_authenticationSettingsModel;

    ExcelTableReaderNodeDialog(final SettingsModelReaderFileChooser settingsModelReaderFileChooser,
        final ExcelMultiTableReadConfig //
        config, final ExcelTableReader tableReader,
        final MultiTableReadFactory<FSPath, ExcelTableReaderConfig, KNIMECellType> readFactory,
        final ProductionPathProvider<KNIMECellType> defaultProductionPathProvider) {
        super(readFactory, defaultProductionPathProvider, true);
        m_settingsModelFilePanel = settingsModelReaderFileChooser;
        m_config = config;
        m_tableReader = tableReader;
        final String[] keyChain =
            Stream.concat(Stream.of("settings"), Arrays.stream(m_settingsModelFilePanel.getKeysForFSLocation()))
                .toArray(String[]::new);
        final FlowVariableModel readFvm = createFlowVariableModel(keyChain, FSLocationVariableType.INSTANCE);
        m_filePanel = new DialogComponentReaderFileChooser(m_settingsModelFilePanel, "excel_reader_writer", readFvm);

        m_authenticationSettingsModel = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
        m_passwordComponent = new DialogComponentAuthentication(m_authenticationSettingsModel, null,
            AuthenticationType.PWD, AuthenticationType.CREDENTIALS, AuthenticationType.NONE);

        final Font italicFont = new Font(m_firstSheetWithDataLabel.getFont().getName(), Font.ITALIC,
            m_firstSheetWithDataLabel.getFont().getSize());
        m_firstSheetWithDataLabel.setFont(italicFont);
        m_columnHeaderNoteLabel.setFont(italicFont);
        m_sheetIdxNoteLabel.setFont(italicFont);
        m_readAreaNoteLabel.setFont(italicFont);
        m_sheetSelectionButtonGroup.add(m_radioButtonFirstSheetWithData);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByName);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByIndex);
        m_buttonGroupFormulaError.add(m_radioButtonInsertErrorPattern);
        m_buttonGroupFormulaError.add(m_radioButtonInsertMissingCell);
        m_buttonGroupSheetArea.add(m_radioButtonReadEntireSheet);
        m_buttonGroupSheetArea.add(m_radioButtonReadPartOfSheet);
        m_buttonGrouprowIDGeneration.add(m_radioButtonGenerateRowIDs);
        m_buttonGrouprowIDGeneration.add(m_radioButtonReadRowIDsFromCol);
        m_emptyColHeaderBtnGrp.add(m_emptyColHeaderIndex);
        m_emptyColHeaderBtnGrp.add(m_emptyColHeaderExcelColName);

        final AnalysisComponentModel analysisComponentModel = new AnalysisComponentModel();
        m_previewModel = new TableReaderPreviewModel(analysisComponentModel);
        m_fileContentPreviewController = new FileContentPreviewController<>(readFactory, analysisComponentModel,
            m_previewModel, this::createItemAccessor);

        addTab("Settings", createGeneralSettingsTab());
        addTab(TRANSFORMATION_TAB, createTransformationTab());
        addTab("Advanced Settings", createAdvancedSettingsTab());
        addTab("Encryption", createEncryptionSettingsTab());
        registerDialogChangeListeners();
        registerPreviewChangeListeners();
        registerTabbedPaneChangeListeners();
    }

    private void registerDialogChangeListeners() {
        m_settingsModelFilePanel.addChangeListener(e -> {
            setReadingMultipleFiles(m_settingsModelFilePanel.getFilterMode() != FilterMode.FILE);
            m_reevaluateFormulas.setEnabled(!m_settingsModelFilePanel.getPath().endsWith(".xlsb"));
        });

        m_radioButtonSheetByName.addChangeListener(l -> m_sheetNameSelection
            .setEnabled(m_radioButtonSheetByName.isEnabled() && m_radioButtonSheetByName.isSelected()));
        m_radioButtonSheetByIndex.addChangeListener(l -> m_sheetIndexSelection
            .setEnabled(m_radioButtonSheetByIndex.isEnabled() && m_radioButtonSheetByIndex.isSelected()));

        m_filePanel.getSettingsModel().getFilterModeModel().addChangeListener(l -> toggleFailOnDifferingCheckBox());

        m_columnHeaderCheckBox.addChangeListener(l -> {
            final boolean isSelected = m_columnHeaderCheckBox.isSelected();
            m_columnHeaderSpinner.setEnabled(isSelected);
            m_columnHeaderNoteLabel.setEnabled(isSelected);
            m_emptyColHeaderPrefix.setEnabled(isSelected);
            m_emptyColHeaderLabel.setEnabled(isSelected);
        });

        m_radioButtonInsertErrorPattern
            .addChangeListener(l -> m_formulaErrorPattern.setEnabled(m_radioButtonInsertErrorPattern.isSelected()));

        m_radioButtonReadPartOfSheet.addChangeListener(l -> setEnablednessReadPartOfSheetFields());

        m_radioButtonReadRowIDsFromCol
            .addChangeListener(l -> m_rowIDColumn.setEnabled(m_radioButtonReadRowIDsFromCol.isSelected()));

        m_limitAnalysisChecker
            .addActionListener(e -> m_limitAnalysisSpinner.setEnabled(m_limitAnalysisChecker.isSelected()));

        m_supportChangingFileSchemas.addActionListener(e -> updateTransformationTabEnabledStatus());
    }

    private void setEnablednessReadPartOfSheetFields() {
        final boolean selected = m_radioButtonReadPartOfSheet.isSelected();
        m_fromColLabel.setEnabled(selected);
        m_fromCol.setEnabled(selected);
        m_toColLabel.setEnabled(selected);
        m_toCol.setEnabled(selected);
        m_andLabel.setEnabled(selected);
        m_fromRowLabel.setEnabled(selected);
        m_fromRow.setEnabled(selected);
        m_toRowLabel.setEnabled(selected);
        m_toRow.setEnabled(selected);
        m_dotLabel.setEnabled(selected);
        m_readAreaNoteLabel.setEnabled(selected);
    }

    private void updateTransformationTabEnabledStatus() {
        setEnabled(!m_supportChangingFileSchemas.isSelected(), TRANSFORMATION_TAB);
    }

    private void configChanged(final boolean updateSheets) {
        final boolean isTablePreviewInForeground = isTablePreviewInForeground();
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
            updatePreviewOrFileContentView(isTablePreviewInForeground);
        } else {
            if (!m_updatingSheetSelection) {
                m_tableReader.setChangeListener(null);
                updatePreviewOrFileContentView(isTablePreviewInForeground);
            }
        }
        if (isTablePreviewInForeground) {
            m_previewConfigChanged = false;
        } else {
            m_fileContentConfigChanged = false;
        }
    }

    private void updatePreviewOrFileContentView(final boolean isTablePreviewInForeground) {
        if (isTablePreviewInForeground) {
            if (m_previewConfigChanged) {
                configChanged();
            }
        } else if (m_fileContentConfigChanged) {
            updateFileContentPreview();
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
        final String text;
        if (m_settingsModelFilePanel.getFilterMode() == FilterMode.FILE) {
            // text has trailing space to not be cut on Windows because of italic font
            text = String.format("(%s) ", ExcelUtils.getFirstSheetWithDataOrFirstIfAllEmpty(sheetNames));
        } else {
            text = "";
        }
        m_firstSheetWithDataLabel.setText(text);
        // set the preferred height to be correctly aligned with the radio button
        m_firstSheetWithDataLabel.setPreferredSize(new Dimension((int)new JLabel(text).getPreferredSize().getWidth(),
            (int)m_radioButtonFirstSheetWithData.getPreferredSize().getHeight()));
        final int selectedIndex = (int)m_sheetIndexSelection.getValue();
        m_sheetIndexSelection.setModel(new SpinnerNumberModel(
            selectedIndex > (sheetNames.size() - 1) ? 0 : selectedIndex, 0, sheetNames.size() - 1, 1));

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
        m_sheetIdxNoteLabel.setPreferredSize(new Dimension((int)m_sheetIdxNoteLabel.getPreferredSize().getWidth(),
            (int)m_sheetIndexSelection.getPreferredSize().getHeight()));
        panel.add(m_sheetIdxNoteLabel, gbcBuilder.incX().setWeightX(1 - smallDouble).build());
        return panel;
    }

    private JPanel createColumnHeaderPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Column header"));
        final Insets insets = new Insets(5, 5, 0, 0);
        final GBCBuilder gbcBuilder = new GBCBuilder(insets).resetPos().setWeightX(1).anchorFirstLineStart();

        panel.add(createColumnNameModePanel(), gbcBuilder.build());
        panel.add(createColumnHeaderRowPanel(), gbcBuilder.incY().build());
        panel.add(createColumnHeaderEmptyPrefixPanel(), gbcBuilder.insetBottom(5).incY().build());

        return panel;
    }

    private JPanel createColumnNameModePanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();

        panel.add(m_emptyColHeaderExcelColName,  gbcBuilder.incX().build());
        panel.add(m_emptyColHeaderIndex,  gbcBuilder.incX().build());

        return panel;
    }

    private JPanel createColumnHeaderRowPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();

        /*
         * This sets the text field of the spinner to a fixed column based size
         * because internally the max value of the spinner is used to determine
         * the size even though you set the preferred size.
         */
        final JComponent editor = m_columnHeaderSpinner.getEditor();
        final JFormattedTextField tf = ((JSpinner.DefaultEditor)editor).getTextField();
        tf.setColumns(6);

        m_columnHeaderNoteLabel
            .setPreferredSize(new Dimension((int)m_columnHeaderNoteLabel.getPreferredSize().getWidth(),
                (int)m_columnHeaderCheckBox.getPreferredSize().getHeight()));
        panel.add(m_columnHeaderCheckBox, gbcBuilder.build());
        panel.add(m_columnHeaderSpinner, gbcBuilder.incX().build());
        panel.add(m_columnHeaderNoteLabel, gbcBuilder.incX().insetLeft(5).setWeightX(1).build());
        return panel;
    }

    private JPanel createColumnHeaderEmptyPrefixPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();

        setHeightToComponentHeight(m_emptyColHeaderLabel, m_emptyColHeaderPrefix);
        setWidthTo(m_emptyColHeaderPrefix, 75);
        panel.add(m_emptyColHeaderLabel, gbcBuilder.insetLeft(5).setWeightX(0).build());
        panel.add(m_emptyColHeaderPrefix, gbcBuilder.incX().build());
        panel.add(Box.createHorizontalBox(), gbcBuilder.incX().setWeightX(1).build());

        return panel;
    }

    private JPanel createRowIDPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Row ID"));
        final GBCBuilder gbcBuilder = new GBCBuilder(new Insets(5, 5, 0, 0)).resetPos().anchorFirstLineStart();

        panel.add(m_radioButtonGenerateRowIDs, gbcBuilder.build());
        panel.add(m_radioButtonReadRowIDsFromCol, gbcBuilder.incX().setInsets(new Insets(5, 5, 5, 0)).build());
        setWidthTo(m_rowIDColumn, 50);
        panel.add(m_rowIDColumn, gbcBuilder.incX().setWeightX(1).build());
        return panel;
    }

    private JPanel createSheetAreaPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Sheet area"));
        final Insets insets = new Insets(5, 5, 0, 0);
        final GBCBuilder gbcBuilder = new GBCBuilder(insets).resetPos().anchorFirstLineStart();

        panel.add(m_radioButtonReadEntireSheet, gbcBuilder.build());
        panel.add(m_radioButtonReadPartOfSheet, gbcBuilder.incX().setWidth(1).build());

        setHeightToComponentHeight(m_fromColLabel, m_radioButtonReadPartOfSheet);
        setHeightToComponentHeight(m_toColLabel, m_radioButtonReadPartOfSheet);
        setHeightToComponentHeight(m_andLabel, m_radioButtonReadPartOfSheet);
        setWidthTo(m_fromCol, 50);
        setWidthTo(m_toCol, 50);
        setWidthTo(m_fromRow, 50);
        setWidthTo(m_toRow, 50);

        panel.add(m_fromColLabel, gbcBuilder.incX().setInsets(new Insets(5, 0, 0, 5)).build());
        panel.add(m_fromCol, gbcBuilder.incX().build());
        panel.add(m_toColLabel, gbcBuilder.incX().build());
        panel.add(m_toCol, gbcBuilder.incX().build());
        panel.add(m_andLabel, gbcBuilder.incX().setWidth(2).setWeightX(1).build());
        m_fromRowLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        m_fromRowLabel.setPreferredSize(new Dimension((int)m_fromColLabel.getPreferredSize().getWidth(),
            (int)m_fromRowLabel.getPreferredSize().getHeight()));
        panel.add(m_fromRowLabel, gbcBuilder.incY().resetX().incX().incX().setWidth(1).setWeightX(0)
            .setInsets(new Insets(5, 0, 5, 5)).build());
        panel.add(m_fromRow, gbcBuilder.incX().build());
        panel.add(m_toRowLabel, gbcBuilder.incX().build());
        panel.add(m_toRow, gbcBuilder.incX().build());
        panel.add(m_dotLabel, gbcBuilder.incX().build());
        panel.add(m_readAreaNoteLabel, gbcBuilder.incX().build());
        return panel;
    }

    static void setHeightToComponentHeight(final JComponent compHeightToSet, final JComponent refComp) {
        compHeightToSet.setPreferredSize(new Dimension((int)compHeightToSet.getPreferredSize().getWidth(),
            (int)refComp.getPreferredSize().getHeight()));
    }

    static void setWidthTo(final JComponent comp, final int width) {
        comp.setPreferredSize(new Dimension(width, (int)comp.getPreferredSize().getHeight()));
    }

    private final JPanel createGeneralSettingsTab() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createFilePanel(), gbcBuilder.build());
        panel.add(createSheetSelectionPanel(), gbcBuilder.incY().build());
        panel.add(createColumnHeaderPanel(), gbcBuilder.incY().build());
        panel.add(createRowIDPanel(), gbcBuilder.incY().build());
        panel.add(createSheetAreaPanel(), gbcBuilder.incY().build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    @Override
    protected JComponent createPreviewComponent() {
        final JTabbedPane tabbedPane = new JTabbedPane();
        final TableReaderPreviewView preview = createPreview();
        preview.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Preview with current settings"));
        tabbedPane.add("Preview", preview);
        tabbedPane.add("File Content", createFileContentPreview());
        m_tabbedPreviewPaneList.add(tabbedPane);
        return tabbedPane;
    }

    private JPanel createAdvancedSettingsTab() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createAdvancedReaderOptionsPanel(), gbcBuilder.build());
        panel.add(createFormulaEvaluationErrorOptionsPanel(), gbcBuilder.incY().build());
        panel.add(createDataRowsSpecLimitPanel(), gbcBuilder.incY().build());
        panel.add(createSpecFailingOptionsPanel(), gbcBuilder.incY().build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createEncryptionSettingsTab() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GBCBuilder gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createEncryptionPanel(), gbcBuilder.build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createAdvancedReaderOptionsPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Reading options"));
        final GBCBuilder gbcBuilder =
            new GBCBuilder(new Insets(5, 5, 0, 5)).resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(m_skipEmptyCols, gbcBuilder.build());
        panel.add(m_skipHiddenCols, gbcBuilder.incY().build());
        panel.add(m_skipEmptyRows, gbcBuilder.incY().build());
        panel.add(m_skipHiddenRows, gbcBuilder.incY().build());
        panel.add(m_use15DigitsPrecision, gbcBuilder.incY().build());
        panel.add(m_replaceEmptyStringsWithMissings, gbcBuilder.incY().build());
        panel.add(m_reevaluateFormulas, gbcBuilder.incY().setInsets(new Insets(5, 5, 5, 5)).build());
        return panel;
    }

    private JPanel createEncryptionPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Password for protected files"));
        final GBCBuilder gbcBuilder =
            new GBCBuilder(new Insets(0, 5, 0, 5)).resetPos().anchorFirstLineStart().fillBoth();
        panel.add(m_passwordComponent.getComponentPanel(), gbcBuilder.build());
        panel.add(Box.createHorizontalBox(), gbcBuilder.incX().setWeightX(1).build());
        return panel;
    }

    private JPanel createDataRowsSpecLimitPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Table specification"));
        final GBCBuilder gbcBuilder =
            new GBCBuilder(new Insets(5, 5, 0, 5)).resetPos().anchorFirstLineStart().fillHorizontal();
        panel.add(m_limitAnalysisChecker, gbcBuilder.build());
        setWidthTo(m_limitAnalysisSpinner, 100);
        panel.add(m_limitAnalysisSpinner, gbcBuilder.incX().setWeightX(1).fillNone().build());
        panel.add(m_supportChangingFileSchemas, gbcBuilder.incY().resetX().setWidth(2).build());
        return panel;
    }

    private JPanel createFormulaEvaluationErrorOptionsPanel() {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Formula error handling"));
        final GBCBuilder gbcBuilder = new GBCBuilder(new Insets(5, 5, 0, 5)).resetPos().anchorFirstLineStart();
        panel.add(m_radioButtonInsertErrorPattern, gbcBuilder.build());
        setWidthTo(m_formulaErrorPattern, 150);
        panel.add(m_formulaErrorPattern, gbcBuilder.incX().setWeightX(1).setInsets(new Insets(7, 0, 0, 5)).build());
        panel.add(m_radioButtonInsertMissingCell,
            gbcBuilder.setWeightX(0).resetX().incY().setInsets(new Insets(5, 5, 5, 5)).build());
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
        m_settingsModelFilePanel.addChangeListener(l -> configRelevantForFileContentChanged(true));
        m_authenticationSettingsModel.addChangeListener(l -> configRelevantForFileContentChanged(true));

        final ActionListener actionListener = l -> configNotRelevantForFileContentChanged();
        m_failOnDifferingSpecs.addActionListener(actionListener);
        m_columnHeaderCheckBox.addActionListener(actionListener);
        m_skipHiddenCols.addActionListener(actionListener);
        m_skipHiddenRows.addActionListener(actionListener);
        m_skipEmptyRows.addActionListener(actionListener);
        m_skipEmptyCols.addActionListener(actionListener);
        m_use15DigitsPrecision.addActionListener(actionListener);
        m_replaceEmptyStringsWithMissings.addActionListener(actionListener);
        m_reevaluateFormulas.addActionListener(actionListener);
        m_radioButtonInsertErrorPattern.addActionListener(actionListener);
        m_radioButtonInsertMissingCell.addActionListener(actionListener);
        m_radioButtonReadEntireSheet.addActionListener(actionListener);
        m_radioButtonReadPartOfSheet.addActionListener(actionListener);
        m_radioButtonGenerateRowIDs.addActionListener(actionListener);
        m_radioButtonReadRowIDsFromCol.addActionListener(actionListener);
        m_supportChangingFileSchemas.addActionListener(actionListener);
        m_emptyColHeaderIndex.addActionListener(actionListener);
        m_emptyColHeaderExcelColName.addActionListener(actionListener);

        final ChangeListener changeListener = l -> configNotRelevantForFileContentChanged();
        m_columnHeaderSpinner.addChangeListener(changeListener);


        final ChangeListener changeListenerFileContent = l -> configRelevantForFileContentChanged(false);
        final ActionListener actionListenerFileContent = l -> configRelevantForFileContentChanged(false);
        m_sheetIndexSelection.addChangeListener(changeListenerFileContent);
        m_radioButtonFirstSheetWithData.addActionListener(actionListenerFileContent);
        m_radioButtonSheetByName.addActionListener(actionListenerFileContent);
        m_radioButtonSheetByIndex.addActionListener(actionListenerFileContent);
        m_sheetNameSelection.addActionListener(actionListenerFileContent);
        m_limitAnalysisChecker.addActionListener(actionListenerFileContent);
        m_limitAnalysisSpinner.addChangeListener(changeListenerFileContent);

        final DocumentListener documentListener = new DocumentListener() {

            @Override
            public void removeUpdate(final DocumentEvent e) {
                configNotRelevantForFileContentChanged();
            }

            @Override
            public void insertUpdate(final DocumentEvent e) {
                configNotRelevantForFileContentChanged();
            }

            @Override
            public void changedUpdate(final DocumentEvent e) {
                configNotRelevantForFileContentChanged();
            }
        };
        m_formulaErrorPattern.getDocument().addDocumentListener(documentListener);
        m_fromCol.getDocument().addDocumentListener(documentListener);
        m_toCol.getDocument().addDocumentListener(documentListener);
        m_fromRow.getDocument().addDocumentListener(documentListener);
        m_toRow.getDocument().addDocumentListener(documentListener);
        m_rowIDColumn.getDocument().addDocumentListener(documentListener);
        m_emptyColHeaderPrefix.getDocument().addDocumentListener(documentListener);
    }

    private void registerTabbedPaneChangeListeners() {
        for (final JTabbedPane tabbedPane : m_tabbedPreviewPaneList) {
            tabbedPane.addChangeListener(l -> {
                // we need to catch notifications of the listeners caused by changes done within this listener
                if (!m_switchTabInTabbedPanes) {
                    m_switchTabInTabbedPanes = true;
                    setSelectedIdxToAllTabbedPanes(tabbedPane.getSelectedIndex());
                    triggerPreviewUpdates();
                    m_switchTabInTabbedPanes = false;
                }
            });
        }
    }

    private void triggerPreviewUpdates() {
        configChanged(false);
    }

    private void setSelectedIdxToAllTabbedPanes(final int selectedIndex) {
        for (final JTabbedPane tabbedPane2 : m_tabbedPreviewPaneList) {
            tabbedPane2.setSelectedIndex(selectedIndex);
        }
    }

    private void configNotRelevantForFileContentChanged() {
        m_previewConfigChanged = true;
        configChanged(false);
    }

    private void configRelevantForFileContentChanged(final boolean updateSheets) {
        m_fileContentConfigChanged = true;
        m_previewConfigChanged = true;
        configChanged(updateSheets);
    }

    private void updateFileContentPreview() {
        if (areIOComponentsDisabled()) {
            m_fileContentPreviewController.setDisabledInRemoteJobViewInfo();
        } else if (!areEventsIgnored()) {
            m_fileContentPreviewController.configChanged(this::getConfig);
        }
    }

    private TableReaderPreviewView createFileContentPreview() {
        final TableReaderPreviewView preview = new TableReaderPreviewView(m_previewModel);
        preview.getTableView().getHeaderTable().setColumnName("Row No.");
        preview.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Content of the selected file(s)"));
        m_fileContentPreviews.add(preview);
        preview.addScrollListener(this::updateFileContentScrolling);
        return preview;
    }

    private void updateFileContentScrolling(final ChangeEvent changeEvent) {
        final TableReaderPreviewView updatedView = (TableReaderPreviewView)changeEvent.getSource();
        for (TableReaderPreviewView preview : m_fileContentPreviews) {
            if (preview != updatedView) {
                preview.updateViewport(updatedView);
            }
        }
    }

    private boolean isTablePreviewInForeground() {
        return m_tabbedPreviewPaneList.isEmpty() || m_tabbedPreviewPaneList.get(0).getSelectedIndex() == 0;
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        super.saveSettingsTo(settings);
        m_filePanel.saveSettingsTo(SettingsUtils.getOrAdd(settings, SettingsUtils.CFG_SETTINGS_TAB));
        saveTableReadSettings();
        saveExcelReadSettings();
        m_config.saveInDialog(settings);
    }

    @Override
    protected ExcelMultiTableReadConfig loadSettings(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        // FIXME: loading should be handled by the config (AP-14460 & AP-14462)
        m_filePanel.loadSettingsFrom(SettingsUtils.getOrEmpty(settings, SettingsUtils.CFG_SETTINGS_TAB), specs);
        m_config.loadInDialog(settings, specs);
        loadTableReadSettings();
        loadExcelSettings();
        m_passwordComponent.loadSettingsFrom(
            SettingsUtils.getOrEmpty(settings, ExcelMultiTableReadConfigSerializer.CFG_ENCRYPTION_SETTINGS_TAB), specs,
            getCredentialsProvider());
        m_fileContentConfigChanged = true;
        ignoreEvents(false);
        updatePreviewOrFileContentView(isTablePreviewInForeground());
        return m_config;
    }

    @Override
    public void refreshPreview(final boolean refreshPreview) {
        final boolean enabled = !areIOComponentsDisabled();
        setPreviewEnabled(enabled);
        if (enabled && refreshPreview && isTablePreviewInForeground()) {
            configChanged();
        }
    }

    /**
     * Fill in the setting values in {@link TableReadConfig} using values from dialog.
     */
    private void saveTableReadSettings() {
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = m_config.getTableReadConfig();
        tableReadConfig.setUseRowIDIdx(m_radioButtonReadRowIDsFromCol.isSelected());
        tableReadConfig.setLimitRowsForSpec(m_limitAnalysisChecker.isSelected());
        tableReadConfig.setMaxRowsForSpec((long)m_limitAnalysisSpinner.getValue());
        tableReadConfig.setRowIDIdx(0);
        tableReadConfig.setSkipEmptyRows(m_skipEmptyRows.isSelected());
        tableReadConfig.setUseColumnHeaderIdx(m_columnHeaderCheckBox.isSelected());
        tableReadConfig.setColumnHeaderIdx((long)m_columnHeaderSpinner.getValue() - 1);
        tableReadConfig.setDecorateRead(false);
        m_config.setFailOnDifferingSpecs(m_failOnDifferingSpecs.isSelected());
        boolean saveSpec = !m_supportChangingFileSchemas.isSelected();
        m_config.setSaveTableSpecConfig(saveSpec);
        m_config.setTableSpecConfig(saveSpec ? getTableSpecConfig() : null);
    }

    /**
     * Fill in the setting values in {@link ExcelTableReaderConfig} using values from dialog.
     */
    private void saveExcelReadSettings() throws InvalidSettingsException {
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
        excelConfig.setSkipHiddenCols(m_skipHiddenCols.isSelected());
        excelConfig.setSkipHiddenRows(m_skipHiddenRows.isSelected());
        m_config.setSkipEmptyColumns(m_skipEmptyCols.isSelected());
        excelConfig.setReplaceEmptyStringsWithMissings(m_replaceEmptyStringsWithMissings.isSelected());
        excelConfig.setReevaluateFormulas(m_reevaluateFormulas.isSelected());
        if (m_radioButtonInsertMissingCell.isSelected()) {
            excelConfig.setFormulaErrorHandling(FormulaErrorHandling.MISSING);
        } else {
            excelConfig.setFormulaErrorHandling(FormulaErrorHandling.PATTERN);
        }
        excelConfig.setErrorPattern(m_formulaErrorPattern.getText());
        if (m_radioButtonReadPartOfSheet.isSelected()) {
            excelConfig.setAreaOfSheetToRead(AreaOfSheetToRead.PARTIAL);
        } else {
            excelConfig.setAreaOfSheetToRead(AreaOfSheetToRead.ENTIRE);
        }

        if (m_radioButtonReadRowIDsFromCol.isSelected()) {
            excelConfig.setRowIdGeneration(RowIDGeneration.COLUMN);
        } else {
            excelConfig.setRowIdGeneration(RowIDGeneration.GENERATE);
        }
        try {
            excelConfig.setRowIDCol(m_rowIDColumn.getText());
            excelConfig.setReadFromCol(m_fromCol.getText());
            excelConfig.setReadToCol(m_toCol.getText());
            ExcelUtils.validateColIndexes(excelConfig.getReadFromCol(), excelConfig.getReadToCol());
            excelConfig.setReadFromRow(m_fromRow.getText());
            excelConfig.setReadToRow(m_toRow.getText());
            ExcelUtils.validateRowIndexes(excelConfig.getReadFromRow(), excelConfig.getReadToRow());
            ExcelUtils.getRowIdColIdx(excelConfig.getRowIDCol());
        } catch (IllegalArgumentException e) { // NOSONAR re-throw as InvalidSettingsException
            throw new InvalidSettingsException(e.getMessage());
        }
        excelConfig.setUseRawSettings(false);

        excelConfig.setCredentialsProvider(getCredentialsProvider());
        excelConfig.setAuthenticationSettingsModel(m_authenticationSettingsModel);

        excelConfig.setEmptyColHeaderPrefix(m_emptyColHeaderPrefix.getText());
        if(m_emptyColHeaderIndex.isSelected()) {
            excelConfig.setColumnNameMode(ColumnNameMode.COL_INDEX);
        } else {
            excelConfig.setColumnNameMode(ColumnNameMode.EXCEL_COL_NAME);
        }

    }

    private void updateFileContentPreviewSettings() {
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig =
            m_fileContentPreviewConfig.getTableReadConfig();
        tableReadConfig.setLimitRowsForSpec(m_limitAnalysisChecker.isSelected());
        tableReadConfig.setMaxRowsForSpec((long)m_limitAnalysisSpinner.getValue());
        final ExcelTableReaderConfig excelConfig = tableReadConfig.getReaderSpecificConfig();
        if (m_radioButtonSheetByName.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.NAME);
        } else if (m_radioButtonSheetByIndex.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.INDEX);
        } else {
            excelConfig.setSheetSelection(SheetSelection.FIRST);
        }
        excelConfig.setSheetName((String)m_sheetNameSelection.getSelectedItem());
        excelConfig.setSheetIdx((int)m_sheetIndexSelection.getValue());
        excelConfig.setCredentialsProvider(getCredentialsProvider());
        excelConfig.setAuthenticationSettingsModel(m_authenticationSettingsModel);
    }

    private static ExcelMultiTableReadConfig createFileContentPreviewSettings() {
        final ExcelMultiTableReadConfig multiTableReadConfig = new ExcelMultiTableReadConfig();
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig =
            multiTableReadConfig.getTableReadConfig();
        tableReadConfig.setDecorateRead(false);
        tableReadConfig.setRowIDIdx(0);
        tableReadConfig.setUseRowIDIdx(true);
        tableReadConfig.setSkipEmptyRows(false);
        tableReadConfig.setUseColumnHeaderIdx(false);
        tableReadConfig.setAllowShortRows(true);
        final ExcelTableReaderConfig excelConfig = tableReadConfig.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(true);
        excelConfig.setSkipHiddenCols(false);
        excelConfig.setSkipHiddenRows(false);
        excelConfig.setReplaceEmptyStringsWithMissings(false);
        excelConfig.setReevaluateFormulas(false);
        excelConfig.setFormulaErrorHandling(FormulaErrorHandling.PATTERN);
        excelConfig.setAreaOfSheetToRead(AreaOfSheetToRead.ENTIRE);
        excelConfig.setRowIdGeneration(RowIDGeneration.GENERATE);
        excelConfig.setUseRawSettings(true);
        return multiTableReadConfig;
    }

    /**
     * Fill in dialog components with {@link TableReadConfig} values.
     */
    private void loadTableReadSettings() {
        m_failOnDifferingSpecs.setSelected(m_config.failOnDifferingSpecs());
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = m_config.getTableReadConfig();
        m_skipEmptyRows.setSelected(tableReadConfig.skipEmptyRows());
        m_skipEmptyCols.setSelected(m_config.skipEmptyColumns());
        m_columnHeaderCheckBox.setSelected(tableReadConfig.useColumnHeaderIdx());
        m_columnHeaderSpinner.setValue(tableReadConfig.getColumnHeaderIdx() + 1);
        m_radioButtonReadRowIDsFromCol.setSelected(tableReadConfig.useRowIDIdx());
        m_limitAnalysisChecker.setSelected(tableReadConfig.limitRowsForSpec());
        m_limitAnalysisSpinner.setValue(tableReadConfig.getMaxRowsForSpec());
        //enable disable spinner
        m_limitAnalysisSpinner.setEnabled(m_limitAnalysisChecker.isSelected());
        m_supportChangingFileSchemas.setSelected(!m_config.saveTableSpecConfig());
        updateTransformationTabEnabledStatus();
    }

    /**
     * Fill in dialog components with {@link ExcelTableReaderConfig} values.
     */
    private void loadExcelSettings() {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        m_use15DigitsPrecision.setSelected(excelConfig.isUse15DigitsPrecision());
        loadExcelSettingsSheetSelection(excelConfig);
        m_sheetNameSelection.addItem(excelConfig.getSheetName());
        m_sheetNameSelection.setSelectedItem(excelConfig.getSheetName());
        m_sheetIndexSelection.setValue(excelConfig.getSheetIdx());
        m_skipHiddenCols.setSelected(excelConfig.isSkipHiddenCols());
        m_skipHiddenRows.setSelected(excelConfig.isSkipHiddenRows());
        m_replaceEmptyStringsWithMissings.setSelected(excelConfig.isReplaceEmptyStringsWithMissings());
        m_reevaluateFormulas.setSelected(excelConfig.isReevaluateFormulas());
        loadExcelSettingsFormulaErrorHandling(excelConfig);
        m_formulaErrorPattern.setText(excelConfig.getErrorPattern());
        loadExcelSettingsAreaOfSheet(excelConfig);
        m_fromCol.setText(excelConfig.getReadFromCol());
        m_toCol.setText(excelConfig.getReadToCol());
        m_fromRow.setText(excelConfig.getReadFromRow());
        m_toRow.setText(excelConfig.getReadToRow());
        setEnablednessReadPartOfSheetFields();
        loadExcelSettingsRowIdGeneration(excelConfig);
        m_rowIDColumn.setText(excelConfig.getRowIDCol());
        m_rowIDColumn.setEnabled(m_radioButtonReadRowIDsFromCol.isSelected());
        loadExcelSettingsColumnNameMode(excelConfig);
        m_emptyColHeaderPrefix.setText(excelConfig.getEmptyColHeaderPrefix());
    }

    /**
     * @param excelConfig the {@link ExcelTableReaderConfig}
     */
    private void loadExcelSettingsColumnNameMode(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getColumnNameMode()) {
            case COL_INDEX:
                m_emptyColHeaderIndex.setSelected(true);
                break;
            case EXCEL_COL_NAME:
            default:
                m_emptyColHeaderExcelColName.setSelected(true);
                break;
        }
    }

    /**
     * Loads the row id generation settings.
     *
     * @param excelConfig the {@link ExcelTableReaderConfig}
     */
    private void loadExcelSettingsRowIdGeneration(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getRowIdGeneration()) {
            case COLUMN:
                m_radioButtonReadRowIDsFromCol.setSelected(true);
                break;
            case GENERATE:
            default:
                m_radioButtonGenerateRowIDs.setSelected(true);
                break;
        }
    }

    /**
     * Loads the area of sheet settings.
     *
     * @param excelConfig the {@link ExcelTableReaderConfig}
     */
    private void loadExcelSettingsAreaOfSheet(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getAreaOfSheetToRead()) {
            case PARTIAL:
                m_radioButtonReadPartOfSheet.setSelected(true);
                break;
            case ENTIRE:
            default:
                m_radioButtonReadEntireSheet.setSelected(true);
                break;
        }
    }

    /**
     * Loads the formula error handling settings.
     *
     * @param excelConfig the {@link ExcelTableReaderConfig}
     */
    private void loadExcelSettingsFormulaErrorHandling(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getFormulaErrorHandling()) {
            case MISSING:
                m_radioButtonInsertMissingCell.setSelected(true);
                break;
            case PATTERN:
            default:
                m_radioButtonInsertErrorPattern.setSelected(true);
                break;
        }
    }

    /**
     * Loads the sheet selection settings.
     *
     * @param excelConfig the {@link ExcelTableReaderConfig}
     */
    private void loadExcelSettingsSheetSelection(final ExcelTableReaderConfig excelConfig) {
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
    }

    @Override
    protected MultiTableReadConfig<ExcelTableReaderConfig, KNIMECellType> getConfig() throws InvalidSettingsException {
        if (isTablePreviewInForeground()) {
            saveTableReadSettings();
            saveExcelReadSettings();
            return m_config;
        } else {
            updateFileContentPreviewSettings();
            return m_fileContentPreviewConfig;
        }
    }

    @Override
    protected ReadPathAccessor createReadPathAccessor() {
        return m_settingsModelFilePanel.createReadPathAccessor();
    }

    @Override
    public void onClose() {
        super.onClose();
        m_fileContentPreviewController.onClose();
        m_filePanel.onClose();
    }
}
