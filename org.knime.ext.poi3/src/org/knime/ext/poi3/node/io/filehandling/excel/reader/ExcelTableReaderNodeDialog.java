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
import java.awt.FlowLayout;
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
import javax.swing.JCheckBox;
import javax.swing.JComponent;
import javax.swing.JFormattedTextField;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DialogComponentAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.ViewUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.DialogUtil;
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
import org.knime.filehandling.core.node.table.reader.dialog.SourceIdentifierColumnPanel;
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

    /*
     * The listener hierarchy goes something like this:
     *
     * ┌────────────────┐
     * │ Input Location │
     * │ (File Chooser) │◀────────────────┐
     * └────────────────┘                 │               ┌─────────────────────────┐
     *          ▲                         └───────────────┼─Reevaluate Formulas     │
     *          │                                       ╱ │ Data Area               │
     * ┌────────────────┐           ┌────────────────┐ ╱  │ Column Names from row#  │
     * │Sheet Selection │           │     Flags      │╱   │ Skip...                 │
     * └────────────────┘           │   (Advanced,   │    │ Encryption              │
     *          ▲                   │Checkboxes, ...)│╲   │ Analysis limit          │
     *          │                   └────────────────┘ ╲  │ Values                  │
     * ┌────────────────┐                    ▲          ╲ │ ...                     │
     * │ Preview Panel  │────────────────────┘           ╲│                         │
     * └────────────────┘                                 └─────────────────────────┘
     *
     */

    private final ExcelMultiTableReadConfig m_fileContentPreviewConfig = createFileContentPreviewSettings();
    private final ExcelMultiTableReadConfig m_config;
    private final ExcelTableReader m_tableReader;

    /*
     * Tab: File and Sheet
     */
    private final DialogComponentReaderFileChooser m_filePanel;
    private final SettingsModelReaderFileChooser m_settingsModelFilePanel;

    private final SheetSelectionComponent m_sheetSelection;

    /*
     * Tab: Data Area
     */
    // Read Area
    private final JRadioButton m_radioButtonReadEntireSheet =
            new JRadioButton(AreaOfSheetToRead.ENTIRE.getText(), true);

    private final JRadioButton m_radioButtonReadPartOfSheet = new JRadioButton(AreaOfSheetToRead.PARTIAL.getText());
    private final ButtonGroup m_buttonGroupSheetArea = new ButtonGroup();
    private final JLabel m_fromColLabel = new JLabel("Column");
    private final JLabel m_toColLabel = new JLabel("to");
    private final JLabel m_fromRowLabel = new JLabel("and row");
    private final JLabel m_toRowLabel = new JLabel("to");
    private final JLabel m_dotLabel = new JLabel(".");
    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_readAreaNoteLabel = new JLabel("(See \"File Content\" tab to identify columns and rows.) ");
    private final JTextField m_fromCol = new JTextField("A");
    private final JTextField m_toCol = new JTextField();
    private final JTextField m_fromRow = new JTextField("1");
    private final JTextField m_toRow = new JTextField();

    // Column Names
    private final JCheckBox m_columnHeaderCheckBox = new JCheckBox("Use values in row", true);
    private final JSpinner m_columnHeaderSpinner = new JSpinner(
        new SpinnerNumberModel(Long.valueOf(1), Long.valueOf(1), Long.valueOf(Long.MAX_VALUE), Long.valueOf(1)));

    // Skip
    private final JCheckBox m_skipHiddenCols = new JCheckBox("Hidden columns", true);
    private final JCheckBox m_skipHiddenRows = new JCheckBox("Hidden rows", true);
    private final JCheckBox m_skipEmptyRows = new JCheckBox("Empty rows", true);
    private final JCheckBox m_skipEmptyCols = new JCheckBox("Empty columns", false);

    /*
     * Tab: Advanced
     */

    /*
     * Panel: File and Sheet
     */
    private final DialogComponentAuthentication m_passwordComponent;
    private final SettingsModelAuthentication m_authenticationSettingsModel;

    private final JCheckBox m_limitAnalysisChecker = new JCheckBox("Limit scanned rows", true);
    private final JSpinner m_limitAnalysisSpinner = new JSpinner(
        new SpinnerNumberModel(Long.valueOf(50), Long.valueOf(1), Long.valueOf(Long.MAX_VALUE), Long.valueOf(50)));

    private final JRadioButton m_failOnChangedSchema = new JRadioButton("Fail");
    private final JRadioButton m_useNewSchema = new JRadioButton("Use new schema");
    private final JRadioButton m_ignoreChangedSchema = new JRadioButton("Ignore (deprecated)");

    private JCheckBox m_failOnDifferingSpecs = new JCheckBox("Fail if schemas differ between multiple files");

    // Checkbox "Append file path column"
    private final SourceIdentifierColumnPanel m_pathColumnPanel = new SourceIdentifierColumnPanel("file path");

    /*
     * Panel: Data Area
     */

    // RowID generation
    private final JRadioButton m_radioButtonGenerateRowIDs = new JRadioButton(RowIDGeneration.GENERATE.getText(), true);
    private final JRadioButton m_radioButtonReadRowIDsFromCol = new JRadioButton(RowIDGeneration.COLUMN.getText());
    private final ButtonGroup m_buttonGrouprowIDGeneration = new ButtonGroup();
    private final JTextField m_rowIDColumn = new JTextField("A");

    // Column name generation (if not provided by row values)
    private final JRadioButton m_emptyColHeaderIndex = new JRadioButton(ColumnNameMode.COL_INDEX.getText());
    private final JRadioButton m_emptyColHeaderExcelColName =
        new JRadioButton(ColumnNameMode.EXCEL_COL_NAME.getText(), true);
    private final ButtonGroup m_emptyColHeaderBtnGrp = new ButtonGroup();

    private final JLabel m_emptyColHeaderLabel = new JLabel("Empty column name prefix:");
    private final JTextField m_emptyColHeaderPrefix = new JTextField("empty_");

    /*
     * Panel: Values
     */
    private final JCheckBox m_reevaluateFormulas =
            new JCheckBox("Reevaluate formulas in all sheets.   On error insert:");
    private final JRadioButton m_radioButtonInsertErrorPattern =
            new JRadioButton(FormulaErrorHandling.PATTERN.getText(), true);
    private final JTextField m_formulaErrorPattern = new JTextField();
    private final JRadioButton m_radioButtonInsertMissingCell =
        new JRadioButton(FormulaErrorHandling.MISSING.getText());
    private final ButtonGroup m_buttonGroupFormulaError = new ButtonGroup();

    private final JCheckBox m_replaceEmptyStringsWithMissings =
            new JCheckBox("Replace empty strings with missing values", true);
    private final JCheckBox m_use15DigitsPrecision = new JCheckBox("Use Excel 15 digits precision");

    /*
     * Misc
     */

    // Preview and File Content
    private List<JTabbedPane> m_tabbedPreviewPaneList = new ArrayList<>();
    private final FileContentPreviewController<ExcelTableReaderConfig, KNIMECellType> m_fileContentPreviewController;
    private final List<TableReaderPreviewView> m_fileContentPreviews = new ArrayList<>();
    private final TableReaderPreviewModel m_previewModel;

    private boolean m_updatingSheetSelection;
    private boolean m_fileContentConfigChanged = true;
    private boolean m_previewConfigChanged;
    private boolean m_switchTabInTabbedPanes;

    private static final String TRANSFORMATION_TAB = "Transformation";

    ExcelTableReaderNodeDialog(final SettingsModelReaderFileChooser settingsModelReaderFileChooser,
        final ExcelMultiTableReadConfig //
        config, final ExcelTableReader tableReader,
        final MultiTableReadFactory<FSPath, ExcelTableReaderConfig, KNIMECellType> readFactory,
        final ProductionPathProvider<KNIMECellType> defaultProductionPathProvider) {
        super(readFactory, defaultProductionPathProvider, true);

        m_settingsModelFilePanel = settingsModelReaderFileChooser;
        m_settingsModelFilePanel.addChangeListener(
            e -> setReadingMultipleFiles(m_settingsModelFilePanel.getFilterMode() != FilterMode.FILE));
        m_settingsModelFilePanel.addChangeListener(
            e -> m_reevaluateFormulas.setEnabled(!m_settingsModelFilePanel.getPath().endsWith(".xlsb")));

        m_config = config;
        m_tableReader = tableReader;

        final String[] keyChain =
            Stream.concat(Stream.of("settings"), Arrays.stream(m_settingsModelFilePanel.getKeysForFSLocation()))
                .toArray(String[]::new);
        final var readFvm = createFlowVariableModel(keyChain, FSLocationVariableType.INSTANCE);
        m_filePanel = new DialogComponentReaderFileChooser(m_settingsModelFilePanel, "excel_reader_writer", readFvm);
        m_filePanel.getSettingsModel().getFilterModeModel().addChangeListener(
            l -> m_failOnDifferingSpecs.setEnabled(m_filePanel.getSettingsModel().getFilterMode() != FilterMode.FILE));

        m_sheetSelection = new SheetSelectionComponent(m_settingsModelFilePanel::getFilterMode);
        m_settingsModelFilePanel.addChangeListener(m_sheetSelection);


        m_columnHeaderCheckBox.addChangeListener(l -> {
            final boolean isSelected = m_columnHeaderCheckBox.isSelected();
            m_columnHeaderSpinner.setEnabled(isSelected);
            m_emptyColHeaderPrefix.setEnabled(isSelected);
            m_emptyColHeaderLabel.setEnabled(isSelected);
        });

        m_authenticationSettingsModel = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
        m_passwordComponent = new DialogComponentAuthentication(m_authenticationSettingsModel, null,
            Arrays.asList(AuthenticationType.PWD, AuthenticationType.CREDENTIALS, AuthenticationType.NONE), Map.of(),
            true);
        m_passwordComponent.setPasswordOnlyLabel("");

        m_limitAnalysisChecker
            .addActionListener(e -> m_limitAnalysisSpinner.setEnabled(m_limitAnalysisChecker.isSelected()));

        final var schemaChangeButtonGroup = new ButtonGroup();
        schemaChangeButtonGroup.add(m_ignoreChangedSchema);
        schemaChangeButtonGroup.add(m_failOnChangedSchema);
        schemaChangeButtonGroup.add(m_useNewSchema);

        m_ignoreChangedSchema.setEnabled(false);
        m_ignoreChangedSchema.addActionListener(e -> updateTransformationTabEnabledStatus());
        m_failOnChangedSchema.addActionListener(e -> updateTransformationTabEnabledStatus());
        m_useNewSchema.addActionListener(e -> updateTransformationTabEnabledStatus());

        DialogUtil.italicizeText(m_readAreaNoteLabel);

        m_radioButtonInsertErrorPattern
            .addChangeListener(l -> m_formulaErrorPattern.setEnabled(m_radioButtonInsertErrorPattern.isSelected()));

        m_radioButtonReadPartOfSheet.addChangeListener(l -> setEnablednessReadPartOfSheetFields());

        m_radioButtonReadRowIDsFromCol
            .addChangeListener(l -> m_rowIDColumn.setEnabled(m_radioButtonReadRowIDsFromCol.isSelected()));

        m_buttonGroupFormulaError.add(m_radioButtonInsertErrorPattern);
        m_buttonGroupFormulaError.add(m_radioButtonInsertMissingCell);
        m_buttonGroupSheetArea.add(m_radioButtonReadEntireSheet);
        m_buttonGroupSheetArea.add(m_radioButtonReadPartOfSheet);
        m_buttonGrouprowIDGeneration.add(m_radioButtonGenerateRowIDs);
        m_buttonGrouprowIDGeneration.add(m_radioButtonReadRowIDsFromCol);
        m_emptyColHeaderBtnGrp.add(m_emptyColHeaderIndex);
        m_emptyColHeaderBtnGrp.add(m_emptyColHeaderExcelColName);


        final var analysisComponentModel = new AnalysisComponentModel();
        m_previewModel = new TableReaderPreviewModel(analysisComponentModel);
        m_fileContentPreviewController = new FileContentPreviewController<>(readFactory, analysisComponentModel,
            m_previewModel, this::createItemAccessor);

        addTab("File and Sheet", createFileAndSheetSettingsTab());
        addTab("Data Area", createDataAreaSettingsTab());
        addTab("Advanced", createAdvancedSettingsTab());
        addTab(TRANSFORMATION_TAB, createTransformationTab());

        registerPreviewChangeListeners();
        registerTabbedPaneChangeListeners();
    }



    private void setEnablednessReadPartOfSheetFields() {
        final boolean selected = m_radioButtonReadPartOfSheet.isSelected();
        m_fromColLabel.setEnabled(selected);
        m_fromCol.setEnabled(selected);
        m_toColLabel.setEnabled(selected);
        m_toCol.setEnabled(selected);
        m_fromRowLabel.setEnabled(selected);
        m_fromRow.setEnabled(selected);
        m_toRowLabel.setEnabled(selected);
        m_toRow.setEnabled(selected);
        m_dotLabel.setEnabled(selected);
        m_readAreaNoteLabel.setEnabled(selected);
    }

    private void updateTransformationTabEnabledStatus() {
        setEnabled(!m_useNewSchema.isSelected(), TRANSFORMATION_TAB);
    }

    private void configChanged(final boolean updateSheets) {
        final boolean isTablePreviewInForeground = isTablePreviewInForeground();
        if (updateSheets) {
            m_updatingSheetSelection = true;
            m_tableReader.setChangeListener(l -> {
                m_tableReader.setChangeListener(null);
                ViewUtils.invokeLaterInEDT(() -> {
                    updateSheetNameSelectionComponent();
                    // update preview again after reconciling actual sheet names/indices
                    updatePreviewOrFileContentView(isTablePreviewInForeground);
                });
            });
            m_sheetSelection.prepareUpdateSheetSelection();
            // update preview immediately to avoid displaying stale data
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

    private void updateSheetNameSelectionComponent() {
        final Map<String, Boolean> sheetNames = m_tableReader.getSheetNames();
        m_sheetSelection.updateSheetSelection(sheetNames);
        m_updatingSheetSelection = false;
    }

    private JPanel createColumnHeaderEmptyPrefixPanel() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();

        setHeightToComponentHeight(m_emptyColHeaderLabel, m_emptyColHeaderPrefix);
        setWidthTo(m_emptyColHeaderPrefix, 75);
        panel.add(m_emptyColHeaderLabel, gbcBuilder.insetLeft(5).setWeightX(0).build());
        panel.add(m_emptyColHeaderPrefix, gbcBuilder.incX().build());
        panel.add(Box.createHorizontalBox(), gbcBuilder.incX().setWeightX(1).build());

        return panel;
    }

    static void setHeightToComponentHeight(final JComponent compHeightToSet, final JComponent refComp) {
        compHeightToSet.setPreferredSize(new Dimension((int)compHeightToSet.getPreferredSize().getWidth(),
            (int)refComp.getPreferredSize().getHeight()));
    }

    static void setWidthTo(final JComponent comp, final int width) {
        comp.setPreferredSize(new Dimension(width, (int)comp.getPreferredSize().getHeight()));
    }

    private JPanel createFileAndSheetSettingsTab() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createFileAndSheetInputLocationPanel(), gbcBuilder.build());
        panel.add(m_sheetSelection.createFileAndSheetSelectSheetPanel("Select Sheet"), gbcBuilder.incY().build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createFileAndSheetInputLocationPanel() {
        final var filePanel = new JPanel();
        filePanel.setLayout(new BoxLayout(filePanel, BoxLayout.X_AXIS));
        filePanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Input Location"));
        filePanel.setMaximumSize(
            new Dimension(Integer.MAX_VALUE, m_filePanel.getComponentPanel().getPreferredSize().height));
        filePanel.add(m_filePanel.getComponentPanel());
        filePanel.add(Box.createHorizontalGlue());
        return filePanel;
    }

    private JPanel createDataAreaSettingsTab() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createDataAreaReadAreaPanel(), gbcBuilder.build());
        panel.add(createDataAreaColumnNamesPanel(), gbcBuilder.incY().build());
        panel.add(createDataAreaSkipPanel(), gbcBuilder.incY().build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createDataAreaReadAreaPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Read Area"));
        final var insets = new Insets(0, 5, 0, 0);
        final var gbcBuilder = new GBCBuilder(insets).resetPos().anchorFirstLineStart();

        panel.add(m_radioButtonReadEntireSheet, gbcBuilder.build());
        panel.add(m_radioButtonReadPartOfSheet, gbcBuilder.incY().build());

        setWidthTo(m_fromCol, 50);
        setWidthTo(m_toCol, 50);
        setWidthTo(m_fromRow, 50);
        setWidthTo(m_toRow, 50);

        final var lableInsets = new Insets(5, 5, 0, 0);
        panel.add(m_fromColLabel, gbcBuilder.setInsets(lableInsets).incX().build());
        panel.add(m_fromCol, gbcBuilder.setInsets(insets).incX().build());
        panel.add(m_toColLabel, gbcBuilder.setInsets(lableInsets).incX().build());
        panel.add(m_toCol, gbcBuilder.setInsets(insets).incX().build());
        panel.add(m_fromRowLabel, gbcBuilder.setInsets(lableInsets).incX().setWeightX(0).build());
        panel.add(m_fromRow, gbcBuilder.setInsets(insets).incX().build());
        panel.add(m_toRowLabel, gbcBuilder.setInsets(lableInsets).incX().build());
        panel.add(m_toRow, gbcBuilder.setInsets(insets).incX().build());
        panel.add(m_dotLabel, gbcBuilder.setInsets(lableInsets).incX().setWeightX(1).fillHorizontal().build());
        return panel;
    }

    private JPanel createDataAreaColumnNamesPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Column Names"));
        final var insets = new Insets(5, 5, 0, 0);
        final var gbcBuilder = new GBCBuilder(insets).resetPos().setWeightX(1).anchorFirstLineStart();
        panel.add(createColumnHeaderRowPanel(), gbcBuilder.build());
        return panel;
    }

    private JPanel createColumnHeaderRowPanel() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();

        /*
         * This sets the text field of the spinner to a fixed column based size
         * because internally the max value of the spinner is used to determine
         * the size even though you set the preferred size.
         */
        final JComponent editor = m_columnHeaderSpinner.getEditor();
        final JFormattedTextField tf = ((JSpinner.DefaultEditor)editor).getTextField();
        tf.setColumns(6);
        panel.add(m_columnHeaderCheckBox, gbcBuilder.build());
        panel.add(m_columnHeaderSpinner, gbcBuilder.incX().build());
        return panel;
    }

    private JPanel createDataAreaSkipPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Skip"));
        final var gbcBuilder = new GBCBuilder(new Insets(5, 5, 0, 5)).resetPos().anchorFirstLineStart();
        panel.add(m_skipEmptyRows, gbcBuilder.build());
        panel.add(m_skipEmptyCols, gbcBuilder.incX().build());
        panel.add(m_skipHiddenRows, gbcBuilder.incX().build());
        panel.add(m_skipHiddenCols, gbcBuilder.incX().setWeightX(1).fillHorizontal().build());
        return panel;
    }

    private JPanel createAdvancedSettingsTab() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).fillHorizontal();
        panel.add(createAdvancedFileAndSheetPanel(), gbcBuilder.build());
        panel.add(createAdvancedDataArePanel(), gbcBuilder.incY().build());
        panel.add(createAdvancedValuesPanel(), gbcBuilder.incY().build());
        panel.add(createPreviewComponent(), gbcBuilder.incY().fillBoth().setWeightY(1).build());
        return panel;
    }

    private JPanel createAdvancedFileAndSheetPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "File and Sheet"));

        final var gbcBuilder = new GBCBuilder(new Insets(5, 5, 0, 5))//
                .resetPos()//
                .anchorFirstLineStart()//
                .fillHorizontal()
                .setWeightX(1);

        panel.add(new JLabel("File encryption"), gbcBuilder.build());
        panel.add(createFileEncryptionPanel(),
            gbcBuilder.incY().insetLeft(20).insetTop(0).build());

        panel.add(new JLabel("Detect columns and data types"),
            gbcBuilder.incY().insetLeft(5).build());
        panel.add(createAdvancedFileSpecPanel(), gbcBuilder.incY().insetLeft(20).insetTop(0).build());

        panel.add(new JLabel("When schema in file has changed"),
            gbcBuilder.incY().insetLeft(5).insetTop(5).build());
        panel.add(createAdvancedSchemaChangedPanel(), gbcBuilder.incY().insetLeft(20).insetTop(0).build());

        //remove the border
        m_pathColumnPanel.setBorder(null);
        panel.add(m_pathColumnPanel, gbcBuilder.incY().insetLeft(5).insetTop(5).build());

        return panel;
    }

    private JPanel createFileEncryptionPanel() {
        final var encryptionPanel = new JPanel(new GridBagLayout());
        final var gbc = new GBCBuilder(new Insets(0, 0, 0, 0)).resetPos().anchorFirstLineStart();

        encryptionPanel.add(m_passwordComponent.getComponentPanel(), gbc.build());
        encryptionPanel.add(Box.createHorizontalGlue(), gbc.incX().setWeightX(1).fillHorizontal().build());

        return encryptionPanel;
    }

    private JPanel createAdvancedFileSpecPanel() {
        final var specPanel = new JPanel(new GridBagLayout());
        final var gbc = new GBCBuilder(new Insets(0, 0, 0, 0)).resetPos().anchorFirstLineStart();

        specPanel.add(m_limitAnalysisChecker, gbc.build());
        setWidthTo(m_limitAnalysisSpinner, 100);
        specPanel.add(m_limitAnalysisSpinner, gbc.incX().build());
        specPanel.add(Box.createHorizontalGlue(), gbc.incX().setWeightX(1).fillHorizontal().build());
        specPanel.add(m_failOnDifferingSpecs, gbc.resetX().incY().insetLeft(0).setWeightX(0).fillNone().build());

        return specPanel;
    }

    private JPanel createAdvancedSchemaChangedPanel() {
        final var specChangedPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));

        specChangedPanel.add(m_failOnChangedSchema);
        specChangedPanel.add(m_useNewSchema);
        specChangedPanel.add(m_ignoreChangedSchema);

        return specChangedPanel;
    }

    private JPanel createAdvancedDataArePanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Data Area"));
        final var lableInsets = new Insets(5, 5, 0, 0);
        final var gbcBuilder = new GBCBuilder(lableInsets).resetPos().anchorFirstLineStart();
        panel.add(new JLabel("RowID"), gbcBuilder.build());
        final var buttonInsets = new Insets(0, 5, 0, 0);
        panel.add(m_radioButtonGenerateRowIDs, gbcBuilder.setInsets(buttonInsets).incX().build());
        panel.add(createRowIDFromColPanel(), gbcBuilder.incX().build());

        panel.add(new JLabel("Column names"), gbcBuilder.setInsets(lableInsets).resetX().incY().build());
        panel.add(m_emptyColHeaderExcelColName, gbcBuilder.setInsets(buttonInsets).incX().build());
        panel.add(m_emptyColHeaderIndex, gbcBuilder.incX().build());
        panel.add(createColumnHeaderEmptyPrefixPanel(), gbcBuilder.setWidth(1).insetLeft(5).incX().build());
        panel.add(new JPanel(), gbcBuilder.setWeightX(1.0).fillHorizontal().incX().build());
        return panel;
    }

    private JPanel createRowIDFromColPanel() {
        final var gbcBuilder = new GBCBuilder().resetPos().anchorFirstLineStart();
        final var rowIdFromCol = new JPanel(new GridBagLayout());
        rowIdFromCol.add(m_radioButtonReadRowIDsFromCol, gbcBuilder.build());
        setWidthTo(m_rowIDColumn, 50);
        rowIdFromCol.add(m_rowIDColumn, gbcBuilder.incX().build());
        return rowIdFromCol;
    }

    private JPanel createAdvancedValuesPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Values"));
        final var gbcBuilder = new GBCBuilder(new Insets(5, 5, 0, 0)).resetPos().anchorFirstLineStart();
        panel.add(m_reevaluateFormulas, gbcBuilder.build());
        panel.add(m_radioButtonInsertErrorPattern, gbcBuilder.incX().setInsets(new Insets(5, 0, 0, 5)).build());
        setWidthTo(m_formulaErrorPattern, 150);
        panel.add(m_formulaErrorPattern, gbcBuilder.incX().build());
        panel.add(m_radioButtonInsertMissingCell, gbcBuilder.incX().setInsets(new Insets(5, 5, 0, 5)).build());
        panel.add(m_replaceEmptyStringsWithMissings, gbcBuilder.resetX().incY().build());
        panel.add(m_use15DigitsPrecision,
            gbcBuilder.incX().setInsets(new Insets(5, 0, 0, 5)).setWidth(3).setWeightX(1).fillHorizontal().build());
        return panel;
    }

    @Override
    protected JComponent createPreviewComponent() {
        final var tabbedPane = new JTabbedPane();
        final TableReaderPreviewView preview = createPreview();
        preview.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Preview with current settings"));
        tabbedPane.add("Preview", preview);
        tabbedPane.add("File Content", createFileContentPreview());
        m_tabbedPreviewPaneList.add(tabbedPane);
        return tabbedPane;
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
        m_ignoreChangedSchema.addActionListener(actionListener);
        m_failOnChangedSchema.addActionListener(actionListener);
        m_useNewSchema.addActionListener(actionListener);
        m_emptyColHeaderIndex.addActionListener(actionListener);
        m_emptyColHeaderExcelColName.addActionListener(actionListener);

        final ChangeListener changeListener = l -> configNotRelevantForFileContentChanged();
        m_columnHeaderSpinner.addChangeListener(changeListener);
        m_pathColumnPanel.addChangeListener(changeListener);

        final ChangeListener changeListenerFileContent = l -> configRelevantForFileContentChanged(false);
        final ActionListener actionListenerFileContent = l -> configRelevantForFileContentChanged(false);

        m_sheetSelection
            .registerConfigRelevantForFileContentChangedListener(l -> configRelevantForFileContentChanged(false));

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
        final var preview = new TableReaderPreviewView(m_previewModel);
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
        m_config.loadInDialog(settings, specs);
        loadTableReadSettings(m_config);
        loadExcelSettings(m_config);
        m_passwordComponent.loadSettingsFrom(
            ExcelMultiTableReadConfigSerializer.getOrDefaultEncryptionSettingsTab(settings), specs,
            getCredentialsProvider());
        // load file after all other settings have loaded to trigger listeners
        // FIXME: loading should be handled by the config (AP-14460 & AP-14462)
        m_filePanel.loadSettingsFrom(SettingsUtils.getOrEmpty(settings, SettingsUtils.CFG_SETTINGS_TAB), specs);
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
        m_config.setAppendItemIdentifierColumn(m_pathColumnPanel.isAppendSourceIdentifierColumn());
        m_config.setItemIdentifierColumnName(m_pathColumnPanel.getSourceIdentifierColumnName());

        // prior to 5.2.2 we had a checkbox "support changing file schemas", which was replaced by a radio button
        // group as part of AP-19239
        boolean saveSpec = !m_useNewSchema.isSelected();
        m_config.setSaveTableSpecConfig(saveSpec);
        m_config.setTableSpecConfig(saveSpec ? getTableSpecConfig() : null);
        m_config.setCheckSavedTableSpec(m_failOnChangedSchema.isSelected());
    }

    /**
     * Fill in the setting values in {@link ExcelTableReaderConfig} using values from dialog.
     */
    private void saveExcelReadSettings() throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(m_use15DigitsPrecision.isSelected());

        m_sheetSelection.saveToExcelConfig(excelConfig);

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
        if (m_emptyColHeaderIndex.isSelected()) {
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
        m_sheetSelection.saveToExcelConfig(excelConfig);
        excelConfig.setCredentialsProvider(getCredentialsProvider());
        excelConfig.setAuthenticationSettingsModel(m_authenticationSettingsModel);
    }

    private static ExcelMultiTableReadConfig createFileContentPreviewSettings() {
        final var multiTableReadConfig = new ExcelMultiTableReadConfig();
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
    private void loadTableReadSettings(final ExcelMultiTableReadConfig config) {
        m_failOnDifferingSpecs.setSelected(config.failOnDifferingSpecs());
        m_pathColumnPanel.load(config.appendItemIdentifierColumn(), config.getItemIdentifierColumnName());
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        m_skipEmptyRows.setSelected(tableReadConfig.skipEmptyRows());
        m_skipEmptyCols.setSelected(config.skipEmptyColumns());
        m_columnHeaderCheckBox.setSelected(tableReadConfig.useColumnHeaderIdx());

        m_columnHeaderSpinner.setValue(tableReadConfig.getColumnHeaderIdx() + 1);
        m_radioButtonReadRowIDsFromCol.setSelected(tableReadConfig.useRowIDIdx());
        m_limitAnalysisChecker.setSelected(tableReadConfig.limitRowsForSpec());
        m_limitAnalysisSpinner.setValue(tableReadConfig.getMaxRowsForSpec());
        //enable disable spinner
        m_limitAnalysisSpinner.setEnabled(m_limitAnalysisChecker.isSelected());

        // prior to 5.2.2 we had a single checkbox "support changing file schemas", which was replaced
        // by a radio button group as part of AP-19239
        if (config.saveTableSpecConfig() && config.checkSavedTableSpec()) {
            m_failOnChangedSchema.setSelected(true);
        } else if (config.saveTableSpecConfig()) {
            m_ignoreChangedSchema.setSelected(true);
        } else {
            m_useNewSchema.setSelected(true);
        }
        updateTransformationTabEnabledStatus();
    }

    /**
     * Fill in dialog components with {@link ExcelTableReaderConfig} values.
     * @param config excel read config
     */
    private void loadExcelSettings(final ExcelMultiTableReadConfig config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        m_use15DigitsPrecision.setSelected(excelConfig.isUse15DigitsPrecision());
        m_sheetSelection.loadExcelSettings(excelConfig);
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
