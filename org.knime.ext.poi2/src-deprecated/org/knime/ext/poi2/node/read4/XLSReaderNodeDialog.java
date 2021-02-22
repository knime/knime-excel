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
package org.knime.ext.poi2.node.read4;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseEvent;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.lang.ref.WeakReference;
import java.nio.file.AccessDeniedException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.ListCellRenderer;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.border.Border;
import javax.swing.border.TitledBorder;
import javax.swing.plaf.basic.BasicComboBoxRenderer;
import javax.swing.table.TableColumnModel;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.lang3.tuple.Triple;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.data.DataTable;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.config.Config;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.tableview.TableContentView;
import org.knime.core.node.tableview.TableContentViewTableHeader;
import org.knime.core.node.tableview.TableView;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.node.util.ViewUtils;
import org.knime.core.node.workflow.FlowVariable.Type;
import org.knime.core.util.MutableInteger;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.ext.poi2.node.read4.POIUtils.StopProcessing;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.defaultnodesettings.DialogComponentFileChooser2;
import org.knime.filehandling.core.defaultnodesettings.FileChooserHelper;
import org.knime.filehandling.core.defaultnodesettings.FileSystemChoice;
import org.knime.filehandling.core.defaultnodesettings.SettingsModelFileChooser2;
import org.knime.filehandling.core.port.FileSystemPortObjectSpec;
import org.xml.sax.SAXException;

/**
 * The dialog to the XLS reader.
 *
 * @author Peter Ohl, KNIME AG, Zurich, Switzerland
 * @author Gabor Bakos
 */
@Deprecated
class XLSReaderNodeDialog extends NodeDialogPane {
    private static final String REFRESH = "refresh";

    private static final String RELOAD = "reload";

    private static final String LOADING_INTERRUPTED = "Loading interrupted";

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSReaderNodeDialog.class);

    private final DialogComponentFileChooser2 m_fileChooser;

    private final JSpinner m_timeout =
        new JSpinner(new SpinnerNumberModel(XLSUserSettings.DEFAULT_TIMEOUT_IN_SECONDS, 0, Integer.MAX_VALUE, 1));

    private final JComboBox<String> m_sheetName = new JComboBox<>();

    private final JCheckBox m_hasColHdr = new JCheckBox();

    private final JTextField m_colHdrRow = new JTextField();

    private final JRadioButton m_hasRowIDs = new JRadioButton();

    private final JRadioButton m_indexContinuous = new JRadioButton();

    private final JRadioButton m_indexSkipJumps = new JRadioButton();

    private final JTextField m_rowIDCol = new JTextField();

    private final JTextField m_firstRow = new JTextField();

    private final JTextField m_lastRow = new JTextField();

    private final JTextField m_firstCol = new JTextField();

    private final JTextField m_lastCol = new JTextField();

    private final JCheckBox m_readAllData = new JCheckBox();

    private final TableView m_fileTable = new TableView();

    private DataTable m_fileDataTable = null;

    private final JPanel m_fileTablePanel = new JPanel();

    private final JPanel m_previewTablePanel = new JPanel();

    private final TableView m_previewTable = new TableView();

    private DataTable m_previewDataTable = null;

    private final JLabel m_previewMsg = new JLabel();

    private final JButton m_previewUpdateButton = new JButton();

    private final JCheckBox m_skipEmptyCols = new JCheckBox();

    private final JCheckBox m_skipHiddenColumns = new JCheckBox("Skip hidden columns");

    private final JCheckBox m_skipEmptyRows = new JCheckBox();

    private final JCheckBox m_uniquifyRowIDs = new JCheckBox();

    private final JRadioButton m_formulaMissCell = new JRadioButton();

    private final JRadioButton m_formulaStringCell = new JRadioButton();

    private final JTextField m_formulaErrPattern = new JTextField();

    private Workbook m_workbook = null;

    private String m_workbookPath = null;

    private static final int LEFT_INDENT = 25;

    /** Flag to temporarily disable listeners during loading of settings. */
    private boolean m_isCurrentlyLoadingNodeSettings = false;

    private static final String SCANNING = "/* scanning... */";

    /** Select the first sheet with data **/
    static final String FIRST_SHEET = "<first sheet with data>";

    /** config key used to store data table spec. */
    static final String XLS_CFG_TABLESPEC = "XLS_DataTableSpec";

    /** config key used to store id of settings used to create table spec. */
    static final String XLS_CFG_ID_FOR_TABLESPEC = "XLS_SettingsForSpecID";

    private String m_fileAccessError = null;

    private static final String PREVIEWBORDER_MSG = "Preview with current settings";

    private final JCheckBox m_reevaluateFormulae =
        new JCheckBox("Reevaluate formulas (leave unchecked if uncertain; see node description for details)");

    private final JCheckBox m_noPreviewChecker =
        new JCheckBox("Disable Preview " + " (does not compute the output table structure)");

    private final Map<Triple<String, String, Boolean>, WeakReference<CachedExcelTable>> m_sheets =
        new ConcurrentHashMap<>();

    private final JButton m_cancel = new JButton("Quick Scan");

    private final JProgressBar m_loadingProgress = new JProgressBar(SwingConstants.HORIZONTAL);

    /** KNIME columns are {@code 0}-based, Excel columns are {@code 1}-based. */
    private final Map<Integer, Integer> m_mapFromKNIMEColumnsToExcel = new HashMap<>();

    private final AtomicReference<Future<CachedExcelTable>> m_currentlyRunningFuture = new AtomicReference<>();

    private final AtomicReference<SwingWorker<?, ?>> m_currentFileWorker = new AtomicReference<>();

    private final AtomicReference<CachedExcelTable> m_currentTable = new AtomicReference<>();

    private final MutableInteger m_readRows = new MutableInteger(0);

    private final AtomicLong m_updateSheetListId = new AtomicLong(0);

    private Optional<FSConnection> m_fsConnection;

    private final SettingsModelFileChooser2 m_fileChooserSettings;

    /**
     *
     */
    XLSReaderNodeDialog() {

        m_fileChooserSettings = XLSReaderNodeModel.getSettingsModelFileChooser();
        final FlowVariableModel fvm = createFlowVariableModel(
            new String[]{m_fileChooserSettings.getConfigName(), SettingsModelFileChooser2.PATH_OR_URL_KEY},
            Type.STRING);

        m_fileChooser = new DialogComponentFileChooser2(0, m_fileChooserSettings, "XLSReader", JFileChooser.OPEN_DIALOG,
            JFileChooser.FILES_AND_DIRECTORIES, fvm);
        final JPanel dlgTab = new JPanel();
        dlgTab.setLayout(new BoxLayout(dlgTab, BoxLayout.Y_AXIS));

        final JComponent fileBox = getFileBox();
        fileBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Select file to read:"));
        dlgTab.add(fileBox);

        final JPanel settingsBox = new JPanel();
        settingsBox.setLayout(new BoxLayout(settingsBox, BoxLayout.Y_AXIS));
        settingsBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Adjust Settings:"));
        settingsBox.add(getSheetAndTimeOutBox());
        settingsBox.add(getColHdrBox());
        settingsBox.add(getRowIDBox());
        settingsBox.add(getAreaBox());
        settingsBox.add(getXLErrBox());
        settingsBox.add(getOptionsBox());
        dlgTab.add(settingsBox);
        dlgTab.add(Box.createVerticalGlue());
        dlgTab.add(getTablesBox());

        addTab("XLS Reader Settings", new JScrollPane(dlgTab));

    }

    private JComponent getFileBox() {
        final JPanel panel = new JPanel(new GridBagLayout());
        final GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 3, 3, 5);
        gbc.weightx = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.gridx = 0;
        gbc.gridy = 0;
        m_fileChooserSettings.addChangeListener(e -> fileNameChanged());
        panel.add(m_fileChooser.getComponentPanel(), gbc);

        return panel;
    }

    @SuppressWarnings("serial")
    private JComponent getSheetAndTimeOutBox() {
        final Box sheetAndTimeOutBox = Box.createHorizontalBox();
        sheetAndTimeOutBox.add(Box.createHorizontalGlue());
        sheetAndTimeOutBox.add(new JLabel("Select the sheet to read:"));
        sheetAndTimeOutBox.add(Box.createHorizontalStrut(5));
        m_sheetName.setPreferredSize(new Dimension(170, 25));
        m_sheetName.setMaximumSize(new Dimension(170, 25));
        m_sheetName.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                if (e.getStateChange() == ItemEvent.SELECTED) {
                    sheetNameChanged();
                }
            }
        });
        final ListCellRenderer<Object> sheetNameRenderer = new BasicComboBoxRenderer() {
            /**
             * {@inheritDoc}
             */
            @Override
            public Component getListCellRendererComponent(@SuppressWarnings("rawtypes") final JList list,
                final Object value, final int index, final boolean isSelected, final boolean cellHasFocus) {
                if ((index > -1) && (value != null)) {
                    list.setToolTipText(value.toString());
                } else {
                    list.setToolTipText(null);
                }
                return super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
            }
        };
        m_sheetName.setRenderer(sheetNameRenderer);
        sheetAndTimeOutBox.add(m_sheetName);
        sheetAndTimeOutBox.add(Box.createHorizontalGlue());
        sheetAndTimeOutBox.add(Box.createHorizontalGlue());

        final JLabel timeoutLabel = new JLabel("Connect timeout [s]: ");
        final String tooltip = "Timeout to connect to the server in seconds";
        timeoutLabel.setToolTipText(tooltip);
        m_timeout.setToolTipText(tooltip);
        m_timeout.addChangeListener(e -> m_fileChooser.setTimeout(readTimeOutInSecondsFromSpinner() * 1000));
        sheetAndTimeOutBox.add(timeoutLabel);
        sheetAndTimeOutBox.add(Box.createHorizontalStrut(5));
        ((JSpinner.DefaultEditor)m_timeout.getEditor()).getTextField().setColumns(4);
        sheetAndTimeOutBox.add(m_timeout);
        sheetAndTimeOutBox.add(Box.createHorizontalGlue());

        return sheetAndTimeOutBox;
    }

    private void sheetNameChanged() {
        m_sheetName.setToolTipText((String)m_sheetName.getSelectedItem());
        updateTimeoutEnabledness();
        // Refresh preview tables
        updateFileTable();

    }

    private JComponent getColHdrBox() {
        final Box colHdrBox = Box.createHorizontalBox();
        colHdrBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Column Names:"));

        m_hasColHdr.setText("Table contains column names in row number:");
        m_hasColHdr.setToolTipText("Enter a number. First row has number 1.");
        m_hasColHdr.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_colHdrRow.setPreferredSize(new Dimension(75, 25));
        m_colHdrRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_colHdrRow);

        colHdrBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        colHdrBox.add(m_hasColHdr);
        colHdrBox.add(Box.createHorizontalStrut(3));
        colHdrBox.add(m_colHdrRow);
        colHdrBox.add(Box.createHorizontalStrut(3));
        colHdrBox.add(new JLabel("(Row numbers start with 1. Mouse over header to see row number.)"));
        colHdrBox.add(Box.createHorizontalGlue());
        return colHdrBox;
    }

    private void checkBoxChanged() {
        m_colHdrRow.setEnabled(m_hasColHdr.isSelected());
        m_rowIDCol.setEnabled(m_hasRowIDs.isSelected());
        m_uniquifyRowIDs.setEnabled(m_hasRowIDs.isSelected());
        m_firstCol.setEnabled(!m_readAllData.isSelected());
        m_lastCol.setEnabled(!m_readAllData.isSelected());
        m_firstRow.setEnabled(!m_readAllData.isSelected());
        m_lastRow.setEnabled(!m_readAllData.isSelected());
        invalidatePreviewTable();
    }

    private JComponent getRowIDBox() {
        final ButtonGroup bg = new ButtonGroup();
        bg.add(m_hasRowIDs);
        bg.add(m_indexContinuous);
        bg.add(m_indexSkipJumps);

        final Box rowBox = Box.createHorizontalBox();
        m_hasRowIDs.setText("Table contains row IDs in column:");
        m_hasRowIDs.setToolTipText("Enter A, B, C, .... or a number 1 ...");
        m_hasRowIDs.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_rowIDCol.setPreferredSize(new Dimension(75, 25));
        m_rowIDCol.setMaximumSize(new Dimension(75, 25));
        m_rowIDCol.setToolTipText("Enter A, B, C, .... or a number 1 ...");
        addFocusLostListener(m_rowIDCol);
        rowBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        rowBox.add(m_hasRowIDs);
        rowBox.add(Box.createHorizontalStrut(3));
        rowBox.add(m_rowIDCol);
        //rowBox.add(Box.createHorizontalGlue());

        m_uniquifyRowIDs.setText("Make row IDs unique");
        m_uniquifyRowIDs.setToolTipText("If checked, row IDs are uniquified "
            + "by adding a suffix if necessary (could cause memory problems with very large data sets).");
        m_uniquifyRowIDs.setSelected(false);
        m_uniquifyRowIDs.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                invalidatePreviewTable();
            }
        });
        rowBox.add(Box.createHorizontalStrut(6));
        rowBox.add(m_uniquifyRowIDs);
        rowBox.add(Box.createHorizontalGlue());

        final Box indexContinousBox = Box.createHorizontalBox();
        m_indexContinuous.setText("Generate RowIDs (index incrementing, starting with 'Row0')");
        m_indexContinuous
            .setToolTipText("The skipped rows (like empty or header rows) do not increase the row id counter");
        m_indexSkipJumps.setText("Generate RowIDs (index as per sheet content, skipped rows will increment index)");
        m_indexSkipJumps
            .setToolTipText("Empty or header rows are not in the result, but keep increasing the row counter");
        m_indexContinuous.addItemListener(e -> checkBoxChanged());
        m_indexSkipJumps.addItemListener(e -> checkBoxChanged());
        indexContinousBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        indexContinousBox.add(m_indexContinuous);
        indexContinousBox.add(Box.createHorizontalStrut(6));
        indexContinousBox.add(m_indexSkipJumps);
        indexContinousBox.add(Box.createHorizontalGlue());

        final Box rowIDBox = Box.createVerticalBox();
        rowIDBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Row IDs:"));
        rowIDBox.add(indexContinousBox);
        rowIDBox.add(rowBox);
        return rowIDBox;
    }

    private JComponent getAreaBox() {

        final Box rowsBox = Box.createHorizontalBox();
        m_firstRow.setPreferredSize(new Dimension(75, 25));
        m_firstRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_firstRow);
        m_lastRow.setPreferredSize(new Dimension(75, 25));
        m_lastRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_lastRow);
        rowsBox.add(Box.createVerticalGlue());
        rowsBox.add(Box.createVerticalGlue());
        rowsBox.add(new JLabel("and read rows from:"));
        rowsBox.add(Box.createHorizontalStrut(3));
        rowsBox.add(m_firstRow);
        rowsBox.add(Box.createHorizontalStrut(3));
        rowsBox.add(new JLabel("to:"));
        rowsBox.add(Box.createHorizontalStrut(3));
        rowsBox.add(m_lastRow);

        final Box colsBox = Box.createHorizontalBox();
        m_firstCol.setPreferredSize(new Dimension(75, 25));
        m_firstCol.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_firstCol);
        m_lastCol.setPreferredSize(new Dimension(75, 25));
        m_lastCol.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_lastCol);
        colsBox.add(Box.createVerticalGlue());
        colsBox.add(Box.createVerticalGlue());
        colsBox.add(new JLabel("read columns from:"));
        colsBox.add(Box.createHorizontalStrut(3));
        colsBox.add(m_firstCol);
        colsBox.add(Box.createHorizontalStrut(3));
        colsBox.add(new JLabel("to:"));
        colsBox.add(Box.createHorizontalStrut(3));
        colsBox.add(m_lastCol);

        m_readAllData.setText("Read entire data sheet, or ...");
        m_readAllData
            .setToolTipText("If checked, cells that contain something (data, format, color, etc.) are read in");
        m_readAllData.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_readAllData.setSelected(false);

        final Box allVBox = Box.createVerticalBox();
        allVBox.add(m_readAllData);
        allVBox.add(Box.createVerticalGlue());
        allVBox.add(Box.createVerticalGlue());

        final Box fromToVBox = Box.createVerticalBox();
        fromToVBox.add(colsBox);
        fromToVBox.add(Box.createVerticalStrut(5));
        fromToVBox.add(rowsBox);
        fromToVBox.add(Box.createVerticalStrut(5));

        final Box mainAreaBox = Box.createHorizontalBox();
        mainAreaBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        mainAreaBox.add(allVBox);
        mainAreaBox.add(Box.createHorizontalStrut(10));
        mainAreaBox.add(fromToVBox);
        mainAreaBox.add(Box.createHorizontalGlue());

        final Box areaBox = Box.createVerticalBox();
        areaBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(),
            "Select the columns and rows to read:"));
        areaBox.add(mainAreaBox);
        areaBox.add(Box.createVerticalStrut(5));
        areaBox.add(ViewUtils.getInFlowLayout(new JLabel(
            "Tip: Mouse over the column and row headers in the \"File Content\" tab to identify cell coordinates")));
        return areaBox;
    }

    private JComponent getOptionsBox() {

        final JComponent skipBox = getSkipEmptyThingsBox();

        final Box optionsBox = Box.createHorizontalBox();
        optionsBox.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "More Options:"));
        optionsBox.add(skipBox);
        optionsBox.add(getEvaluationBox());
        optionsBox.add(Box.createHorizontalGlue());
        return optionsBox;

    }

    /**
     * @return
     */
    private JComponent getEvaluationBox() {
        final Box evaluationBox = Box.createVerticalBox();
        evaluationBox.add(m_reevaluateFormulae);
        m_reevaluateFormulae.setToolTipText(
            "When checked not the cached values, but the reevaluated values are returned (using DOM representation, "
                + "requires significantly more memory, for xls files, it always uses the DOM representation)");
        m_reevaluateFormulae.addActionListener(e -> sheetNameChanged());
        evaluationBox.add(m_noPreviewChecker);
        m_noPreviewChecker.addItemListener(e -> onNoPreviewCheckerSelected());
        evaluationBox.add(Box.createHorizontalGlue());
        return evaluationBox;
    }

    private JComponent getSkipEmptyThingsBox() {
        final Box skipColsBox = Box.createHorizontalBox();
        m_skipEmptyCols.setText("Skip empty columns");
        m_skipEmptyCols
            .setToolTipText("If checked, columns that contain only missing values are not part of the output table");
        m_skipEmptyCols.setSelected(true);
        m_skipEmptyCols.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                invalidatePreviewTable();
            }
        });
        skipColsBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        skipColsBox.add(m_skipEmptyCols);
        skipColsBox.add(Box.createHorizontalGlue());

        final Box skipHiddenColumns = Box.createHorizontalBox();
        m_skipHiddenColumns.setToolTipText("If checked, hidden column's content is not included in the output");
        m_skipHiddenColumns.setSelected(true);
        m_skipHiddenColumns.addItemListener(e -> invalidatePreviewTable());
        skipHiddenColumns.add(Box.createHorizontalStrut(LEFT_INDENT));
        skipHiddenColumns.add(m_skipHiddenColumns);
        skipHiddenColumns.add(Box.createHorizontalGlue());

        final Box skipRowsBox = Box.createHorizontalBox();
        m_skipEmptyRows.setText("Skip empty rows");
        m_skipEmptyRows
            .setToolTipText("If checked, rows that contain only missing values are not part of the output table");
        m_skipEmptyRows.setSelected(true);
        m_skipEmptyRows.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                invalidatePreviewTable();
            }
        });
        skipRowsBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        skipRowsBox.add(m_skipEmptyRows);
        skipRowsBox.add(Box.createHorizontalGlue());

        final Box skipBox = Box.createVerticalBox();
        skipBox.add(skipColsBox);
        skipBox.add(skipHiddenColumns);
        skipBox.add(skipRowsBox);
        skipBox.add(Box.createVerticalGlue());
        return skipBox;
    }

    private JComponent getXLErrBox() {
        m_formulaMissCell.setText("Insert a missing cell");
        m_formulaMissCell.setToolTipText("A missing cell doesn't change the column's type, but might be hard to spot");
        m_formulaStringCell.setText("Insert an error pattern:");
        m_formulaStringCell.setToolTipText("When the evaluation fails the column becomes a string column");
        final ButtonGroup bg = new ButtonGroup();
        bg.add(m_formulaMissCell);
        bg.add(m_formulaStringCell);
        m_formulaStringCell.setSelected(true);
        m_formulaErrPattern.setColumns(15);
        m_formulaErrPattern.setText(XLSUserSettings.DEFAULT_ERR_PATTERN);
        addFocusLostListener(m_formulaErrPattern);
        m_formulaStringCell.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(final ItemEvent e) {
                m_formulaErrPattern.setEnabled(m_formulaStringCell.isSelected());
                invalidatePreviewTable();
            }
        });

        final JPanel missingBox = new JPanel(new FlowLayout(FlowLayout.LEFT));
        missingBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        missingBox.add(m_formulaMissCell);

        final JPanel stringBox = new JPanel(new FlowLayout(FlowLayout.LEFT));
        stringBox.add(Box.createHorizontalStrut(LEFT_INDENT));
        stringBox.add(m_formulaStringCell);
        stringBox.add(m_formulaErrPattern);

        final Box formulaErrBox = Box.createVerticalBox();
        formulaErrBox
            .setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "On evaluation error:"));
        formulaErrBox.add(stringBox);
        formulaErrBox.add(missingBox);
        formulaErrBox.add(Box.createVerticalGlue());
        return formulaErrBox;
    }

    private void updateTimeoutEnabledness() {
        // Enable/Disable timeout spinner
        final SettingsModelFileChooser2 model = m_fileChooserSettings.clone();
        final FileSystemChoice choice = model.getFileSystemChoice();
        m_timeout.setEnabled(FileSystemChoice.getCustomFsUrlChoice().equals(choice)
            || FileSystemChoice.getKnimeFsChoice().equals(choice));
    }

    private void fileNameChanged() {
    	updateTimeoutEnabledness();

    	// Refresh the workbook when the selected file changed
        refreshWorkbook(m_fileChooserSettings.getPathOrURL());
        // refresh workbook sets the workbook null in case of an error
        if (m_isCurrentlyLoadingNodeSettings) {
            return;
        }
        clearTableViews();
        updateSheetListAndSelect(null);
    }

    /**
     * Reads from the currently selected file the list of worksheets (in a background thread) and selects the provided
     * sheet (if not null - otherwise selects the first name). Calls {@link #sheetNameChanged()} after the update.
     *
     * @param sheetName
     */
    private void updateSheetListAndSelect(final String sheetName) {
        m_previewUpdateButton.setEnabled(false);
        m_previewMsg.setText("Loading input file...");
        m_sheetName.setModel(new DefaultComboBoxModel<>(new String[]{SCANNING}));
        // The id of the current update
        // Note that this code and the doneWithContext is always executed by the same thread
        // Therefore we only have to make sure that the doneWithContext belongs to the most current update
        final long currentId = m_updateSheetListId.incrementAndGet();
        final SwingWorker<String[], Object> sw = new SwingWorkerWithContext<String[], Object>() {

            @Override
            protected String[] doInBackgroundWithContext() throws Exception {
                final List<Path> paths = getFileChooserHelper().getPaths();

                if (paths != null && !paths.isEmpty()) {
                    final Path path = paths.get(0);

                    if ((m_workbook == null) && !ExcelTableReader.isXlsx(path)) {
                        m_workbook = getWorkbook(path);
                    }
                    if (m_workbook != null) {
                        m_fileAccessError = null;
                        final ArrayList<String> sheetNames = POIUtils.getSheetNames(m_workbook);
                        sheetNames.add(0, FIRST_SHEET);
                        return sheetNames.toArray(new String[sheetNames.size()]);
                    } else {//xlsx without reevaluation
                        try (final InputStream stream = Files.newInputStream(path);
                                final OPCPackage opcpackage = OPCPackage.open(stream)) {
                            final List<String> sheetNames = POIUtils.getSheetNames(new XSSFReader(opcpackage));
                            sheetNames.add(0, FIRST_SHEET);
                            return sheetNames.stream().toArray(n -> new String[n]);
                        }
                    }
                } else {
                    m_previewMsg.setText("No input file available");
                }
                return new String[]{};
            }

            /**
             * {@inheritDoc}
             */
            @Override
            protected void doneWithContext() {
                if (currentId != m_updateSheetListId.get()) {
                    // Another update of the sheet list has started
                    // Do not update the sheet list
                    return;
                }
                String[] names = new String[]{};
                try {
                    names = get();
                } catch (InterruptedException e) {
                    fixInterrupt();
                    // ignore
                } catch(ExecutionException e) {
                    if (e.getCause() instanceof NoSuchFileException) {
                        m_fileAccessError = "No such file";
                    } else if (e.getCause() instanceof AccessDeniedException) {
                        m_fileAccessError = String.format("Access was denied for %s", ((AccessDeniedException)e.getCause()).getFile());
                    } else {
                        m_fileAccessError = e.getCause().getMessage();
                    }
                }

                m_sheetName.setModel(new DefaultComboBoxModel<>(names));
                if (names.length > 0) {
                    if (sheetName != null) {
                        m_sheetName.setSelectedItem(sheetName);
                    } else {
                        m_sheetName.setSelectedIndex(0);
                    }
                } else {
                    m_sheetName.setSelectedIndex(-1);
                }
                sheetNameChanged();

            }
        };
        sw.execute();
    }

    private final FileChooserHelper getFileChooserHelper() throws IOException {
        // timeout is passed from JSpinner, but is only used if Custom or KNIME file system is used
        return new FileChooserHelper(m_fsConnection, m_fileChooserSettings.clone(),
            (int)m_timeout.getValue() * 1000);
    }

    private JComponent getTablesBox() {

        final JTabbedPane viewTabs = new JTabbedPane();

        m_fileTablePanel.setLayout(new BorderLayout());
        m_fileTablePanel
            .setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "XL Sheet Content:"));
        m_fileTablePanel.add(m_fileTable, BorderLayout.CENTER);
        m_cancel.setVisible(false);
        m_loadingProgress.setVisible(false);
        m_cancel.addActionListener(e -> onCancelButtonClicked());
        m_fileTable.getHeaderTable().setColumnName("Row No.");
        m_previewTablePanel.setLayout(new BorderLayout());
        m_previewTablePanel
            .setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), PREVIEWBORDER_MSG));
        m_previewTablePanel.add(m_previewTable, BorderLayout.CENTER);
        m_previewUpdateButton.setText(REFRESH);
        m_previewUpdateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(final ActionEvent e) {
                m_previewUpdateButton.setEnabled(false);
                if (RELOAD.equals(m_previewUpdateButton.getText())) {
                    updateFileTable();
                } else {
                    updatePreviewTable();
                }
            }
        });
        m_previewTable.getContentTable()
            .setTableHeader(new TableContentViewTableHeaderWithTooltip(m_previewTable.getContentTable(),
                m_previewTable.getContentTable().getColumnModel(), m_mapFromKNIMEColumnsToExcel));
        m_previewMsg.setForeground(Color.RED);
        m_previewMsg.setText("");
        final Box errBox = Box.createHorizontalBox();
        errBox.add(m_previewUpdateButton);
        errBox.add(Box.createHorizontalStrut(5));
        errBox.add(m_previewMsg);
        errBox.add(Box.createHorizontalGlue());
        errBox.add(Box.createRigidArea(new Dimension(5, 0)));
        errBox.add(m_loadingProgress);
        errBox.add(Box.createRigidArea(new Dimension(5, 0)));
        errBox.add(m_cancel);
        m_cancel.setToolTipText("Stops further loading of the preview, specify the columns based on the read rows.");
        errBox.add(Box.createVerticalStrut(30));
        m_previewTablePanel.add(errBox, BorderLayout.NORTH);
        viewTabs.addTab("Preview", m_previewTablePanel);
        viewTabs.addTab("File Content", m_fileTablePanel);

        return viewTabs;
    }

    private void clearTableViews() {
        ViewUtils.invokeLaterInEDT(() -> {
            setNewPreviewDataTableInEDT(null);
            setNewFileDataTableInEDT(null);
        });
    }

    private void onCancelButtonClicked() {
        m_cancel.setEnabled(false);
        m_previewMsg.setText(interruptedMessage());
        checkPreviousFuture();
    }

    private void onNoPreviewCheckerSelected() {
        final boolean noPreview = m_noPreviewChecker.isSelected();
        m_fileTable.setEnabled(!noPreview);
        m_previewTable.setEnabled(!noPreview);
        m_previewUpdateButton.setEnabled(!noPreview);
        m_previewTablePanel.setEnabled(!noPreview);
        m_fileTablePanel.setEnabled(!noPreview);
        if (noPreview) {
            m_previewUpdateButton.setText(RELOAD);
            m_previewMsg.setText("");
            checkPreviousFuture();
            clearTableViews();
        }
    }

    /**
     * Updates the file TableView and also sets the field.
     *
     * @param newFileDataTable The new file table to set or null
     */
    private void setNewFileDataTableInEDT(final DataTable newFileDataTable) {
        assert SwingUtilities.isEventDispatchThread() : "Not run in EDT";
        if (m_fileDataTable instanceof Closeable) {
            try {
                ((Closeable)m_fileDataTable).close();
            } catch (final IOException e) {
                throw new UncheckedIOException(e);
            }
        }
        m_fileDataTable = newFileDataTable;
        m_fileTable.setDataTable(newFileDataTable);
    }

    /**
     * Updates the preview TableView and also sets the field.
     *
     * @param newPreviewDataTable The new file table to set or null
     */
    private void setNewPreviewDataTableInEDT(final DataTable newPreviewDataTable) {
        assert SwingUtilities.isEventDispatchThread() : "Not run in EDT";
        if (m_previewDataTable instanceof Closeable) {
            try {
                ((Closeable)m_previewDataTable).close();
            } catch (final IOException e) {
                throw new UncheckedIOException(e);
            }
        }
        m_previewDataTable = newPreviewDataTable;
        m_previewTable.setDataTable(newPreviewDataTable);
    }

    /**
     * reads the current filename and sheetname and fills the file content view.
     */
    private void updateFileTable() {
        if (m_noPreviewChecker.isSelected()) {
            m_previewUpdateButton.setText(RELOAD);
            m_previewMsg.setText("");
            return;
        }
        m_previewUpdateButton.setText(REFRESH);

        final String sheet = (String)m_sheetName.getSelectedItem();
        if ((sheet == null) || sheet.isEmpty()) {
            fileNotFound();
            return;
        }
        if (sheet == SCANNING) {
            setFileTablePanelBorderTitle("Loading input file...");
            clearTableViews();
            return;
        }
        if (m_isCurrentlyLoadingNodeSettings) {
            // do not read from the file while loading settings.
            return;
        }
        //To prevent showing invalid content:
        clearTableViews();
        checkPreviousFuture();
        m_currentTable.set(null);
        final AtomicReference<DataTable> dt = new AtomicReference<>(null);
        m_readRows.setValue(-1);
        final SwingWorker<String, Object> sw = new SwingWorkerWithContext<String, Object>() {
            @Override
            protected String doInBackgroundWithContext() throws Exception {
                try {
                    final Path path;
                    try {
                        final List<Path> paths = getFileChooserHelper().getPaths();
                        if (paths.isEmpty()) {
                            return "<no file set>";
                        }
                        path = paths.get(0);
                    } catch (final Exception e1) {
                        return "<could not get file(s)>";
                    }

                    String localSheet = sheet;
                    if (localSheet.equals(FIRST_SHEET)) {
                        localSheet = firstSheetName(path);
                        if (localSheet == null) {
                            return "<could not load the file>";
                        }
                    }
                    final XLSUserSettings settings = createSettingsFromComponents();
                    try {
                        setPreviewTablePanelBorderTitle("Preview with current settings: [" + localSheet + "]");
                        setFileTablePanelBorderTitle("Content of xls(x) sheet: " + localSheet);
                        CachedExcelTable table = getSheetTable(path, localSheet, settings.isReevaluateFormulae());
                        fileTableSettings(localSheet, settings);
                        if (table != null) {
                            m_currentTable.set(table);
                        } else {
                            //Wait for the incomplete results.
                            Thread.sleep(100);
                            table = m_currentTable.get();
                        }
                        if (table == null) {
                            return "<could not load the table>";
                        }
                        m_readRows.setValue(table.lastRow());
                        dt.set(table.createDataTable(settings, null));
                        if (m_readRows.intValue() < 0) {
                            m_previewMsg.setText(interruptedMessage());
                        }
                        return (m_readRows.intValue() >= 0 ? "Content of xls(x) sheet: "
                            : "The first " + -(m_readRows.intValue() + 1) + " values of the sheet: ") + localSheet;
                    } catch (InterruptedException | CancellationException e) {
                        fixInterrupt();
                        checkPreviousFuture();
                        m_previewUpdateButton.setText(RELOAD);
                        final CachedExcelTable cachedExcelTable = m_currentTable.get();
                        if (cachedExcelTable != null) {
                            fileTableSettings(localSheet, settings);
                            dt.set(cachedExcelTable.createDataTable(settings, null));
                        }
                        cancel(false);
                        return "<load was interrupted>";
                    }
                } catch (final Throwable t) {
                    checkPreviousFuture();
                    m_previewUpdateButton.setText(RELOAD);
                    NodeLogger.getLogger(XLSReaderNodeDialog.class)
                        .debug("Unable to create settings for file content view", t);
                    clearTableViews();
                    return "<unable to create view>";
                } finally {
                    m_loadingProgress.setVisible(false);
                    m_cancel.setVisible(false);
                }
            }

            /**
             * @param localSheet
             * @param settings
             */
            private void fileTableSettings(final String localSheet, final XLSUserSettings settings) {
                settings.setHasColHeaders(false);
                settings.setHasRowHeaders(false);
                settings.setKeepXLNames(true);
                settings.setReadAllData(true);
                settings.setFirstRow(0);
                settings.setLastRow(0);
                settings.setFirstColumn(0);
                settings.setLastColumn(0);
                settings.setReevaluateFormulae(false);
                settings.setSkipHiddenColumns(true);
                settings.setSkipEmptyRows(false);
                settings.setSkipHiddenColumns(false);
                settings.setSheetName(localSheet);
            }

            /**
             * {@inheritDoc}
             */
            @Override
            protected void doneWithContext() {
                String msg;
                try {
                    msg = get();
                } catch (InterruptedException | ExecutionException | CancellationException e) {
                    fixInterrupt();
                    msg = "<unable to create view>";
                }
                setFileTablePanelBorderTitle(msg);
                final DataTable newFileDataTable = dt.get();
                setNewFileDataTableInEDT(newFileDataTable);
                m_currentFileWorker.set(null);
                updatePreviewTable();
            }

        };
        setFileTablePanelBorderTitle("Updating file content view...");
        m_previewMsg.setText("Loading input file...");
        final SwingWorker<?, ?> oldWorker = m_currentFileWorker.getAndSet(sw);
        if (oldWorker != null) {
            oldWorker.cancel(true);
        }
        sw.execute();
    }

    void fileNotFound() {
        String msg = "Could not load file";
        setFileTablePanelBorderTitle(msg);
        setPreviewTablePanelBorderTitle(msg);
        clearTableViews();
        if (m_fileAccessError != null) {
            msg += ": " + m_fileAccessError;
        }
        m_previewMsg.setText(msg);
    }

    /**
     * Should only be called from a background thread as it might load a file.
     *
     * @param file File name.
     * @param sheet Sheet name.
     * @param reevaluateFormulae Reevaluate formulae or not?
     * @return The {@link CachedExcelTable}.
     */
    private CachedExcelTable getSheetTable(final Path path, final String sheet, final boolean reevaluateFormulae) {
        CachedExcelTable sheetTable;
        final Triple<String, String, Boolean> key = Triple.of(path.toString(), sheet, reevaluateFormulae);
        if (!m_sheets.containsKey(key) || ((sheetTable = m_sheets.get(key).get()) == null)) {
            LOGGER.debug("Loading sheet " + sheet + "  of " + path.getFileName().toString());

            try (InputStream stream = Files.newInputStream(path)) {
                checkPreviousFuture();
                final ExecutionMonitor monitor = new ExecutionMonitor();
                monitor.getProgressMonitor().addProgressListener(
                    e -> m_loadingProgress.setValue((int)(100 * e.getNodeProgress().getProgress())));
                final Future<CachedExcelTable> tableFuture = ExcelTableReader.isXlsx(path) && !reevaluateFormulae
                    ? CachedExcelTable.fillCacheFromXlsxStreaming(path, stream, sheet, Locale.ROOT, monitor,
                        m_currentTable)
                    : CachedExcelTable.fillCacheFromDOM(path, stream, sheet, Locale.ROOT, reevaluateFormulae,
                        monitor, m_currentTable);
                checkPreviousFutureAndCancel(m_currentlyRunningFuture.getAndSet(tableFuture));
                ViewUtils.invokeAndWaitInEDT(() -> {
                    m_loadingProgress.setValue(0);
                    m_loadingProgress.setVisible(true);
                    m_cancel.setEnabled(true);
                    m_cancel.setVisible(true);
                });
                sheetTable = tableFuture.get();
                if (!m_currentlyRunningFuture.compareAndSet(tableFuture, null)) {
                    LOGGER.warn("Inconsistency, another thread changed the running future");
                    checkPreviousFuture();
                    m_currentlyRunningFuture.set(null);
                }
                m_sheets.put(key, new WeakReference<>(sheetTable));
            } catch (CancellationException | StopProcessing | InterruptedException | ExecutionException e) {
                sheetTable = m_currentTable.get();
                m_previewUpdateButton.setText(RELOAD);
                fixInterrupt();
                m_currentlyRunningFuture.set(null);
            } catch (final IOException e) {
                throw new RuntimeException(e);
            }
        }
        m_currentTable.set(sheetTable);
        return sheetTable;
    }

    /** @return int value from {@link #m_timeout}. */
    private int readTimeOutInSecondsFromSpinner() {
        return ((Number)m_timeout.getValue()).intValue();
    }

    /**
     *
     */
    private static void fixInterrupt() {
        final Thread currentThread = Thread.currentThread();
        if (currentThread.isInterrupted()) {
            currentThread.interrupt();
        }
    }

    /**
     * Checks whether previous future exists and if not finished yet, it cancels.
     */
    private boolean checkPreviousFuture() {
        final Future<CachedExcelTable> previousFuture = m_currentlyRunningFuture.get();
        return checkPreviousFutureAndCancel(previousFuture);
    }

    /**
     * @param previousFuture The previous {@link Future} to cancel.
     */
    private static boolean checkPreviousFutureAndCancel(final Future<CachedExcelTable> previousFuture) {
        if ((previousFuture != null) && !previousFuture.isDone()) {
            LOGGER.debug("Cancelling loading");
            return previousFuture.cancel(true);
        }
        return false;
    }

    /**
     * Should only be called from a background thread, not on EDT.
     *
     * @param path The file path.
     * @return The name of the first sheet with data.
     */
    String firstSheetName(final Path path) {
        if (ExcelTableReader.isXlsx(path)) {
            try (final InputStream stream = Files.newInputStream(path);
                    final OPCPackage opcpackage = OPCPackage.open(stream)) {
                final XSSFReader reader = new XSSFReader(opcpackage);
                return POIUtils.getFirstSheetNameWithData(reader, new ReadOnlySharedStringsTable(opcpackage));
            } catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException e) {
                return null;
            }
        } else {
            if (m_workbook == null) {
                try {
                    m_workbook = getWorkbook(path);
                } catch (final IOException e) {
                    throw new UncheckedIOException(e);
                } catch (final InvalidFormatException e) {
                    throw new RuntimeException(e);
                }
            }
            return POIUtils.getFirstSheetNameWithData(m_workbook);
        }
    }

    /**
     * Loads a workbook from the file system.
     *
     * @param path Path to the workbook
     * @return The workbook or null if it could not be loaded
     * @throws IOException
     * @throws InvalidFormatException
     * @throws RuntimeException the underlying POI library also throws other kind of exceptions
     */
    public Workbook getWorkbook(final Path path) throws IOException, InvalidFormatException {
        try (InputStream in = Files.newInputStream(path)) {
            // This should be the only place in the code where a workbook gets loaded
            return WorkbookFactory.create(in);
        }
    }

    private void setFileTablePanelBorderTitle(final String title) {
        ViewUtils.invokeAndWaitInEDT(new Runnable() {
            @Override
            public void run() {
                final Border b = m_fileTablePanel.getBorder();
                if (b instanceof TitledBorder) {
                    final TitledBorder tb = (TitledBorder)b;
                    tb.setTitle(title);
                    m_fileTablePanel.repaint();
                }
            }
        });
    }

    private void setPreviewTablePanelBorderTitle(final String title) {
        ViewUtils.invokeAndWaitInEDT(new Runnable() {
            @Override
            public void run() {
                final Border b = m_previewTablePanel.getBorder();
                if (b instanceof TitledBorder) {
                    final TitledBorder tb = (TitledBorder)b;
                    tb.setTitle(title);
                    m_previewTablePanel.repaint();
                }
            }
        });
    }

    private void invalidatePreviewTable() {
        final String txt = m_noPreviewChecker.isSelected() ? ""
            : "Preview table is out of sync with current settings. Please refresh.";
        m_previewMsg.setText(txt);
    }

    /**
     * Call in EDT.
     */
    private void updatePreviewTable() {
        // make sure user doesn't trigger it again
        if (m_isCurrentlyLoadingNodeSettings) {
            // do nothing while loading settings.
            return;
        }

        m_previewUpdateButton.setEnabled(false);
        m_previewMsg.setText("Refreshing preview table....");

        final AtomicReference<DataTable> dt = new AtomicReference<>(null);

        m_readRows.setValue(-1);
        final SwingWorker<String, Object> sw = new SwingWorkerWithContext<String, Object>() {
            @Override
            protected String doInBackgroundWithContext() throws Exception {

                XLSUserSettings s;
                try {
                    s = createSettingsFromComponents();
                    final CachedExcelTable sheetTable = m_currentTable.get();
                    m_mapFromKNIMEColumnsToExcel.clear();
                    if (sheetTable != null) {
                        m_readRows.setValue(sheetTable.lastRow());
                        dt.set(sheetTable.createDataTable(s, m_mapFromKNIMEColumnsToExcel));
                    }
                } catch (final Throwable t) {
                    String msg = t.getMessage();
                    if ((msg == null) || msg.isEmpty()) {
                        msg = "no details, sorry.";
                    }
                    return msg;
                }
                return null;
            }

            /**
             * {@inheritDoc}
             */
            @Override
            protected void doneWithContext() {
                try {
                    setPreviewTablePanelBorderTitle(PREVIEWBORDER_MSG);
                    String err = null;
                    try {
                        err = get();
                    } catch (InterruptedException | ExecutionException e) {
                        err = e.getMessage();
                        fixInterrupt();
                    }
                    if (err != null) {
                        m_previewMsg.setText(err);
                        setNewPreviewDataTableInEDT(null);
                        return;
                    }

                    final String messagePrefix = interruptedMessage();
                    m_previewMsg.setText(messagePrefix);
                    try {
                        final DataTable newPreviewDataTable = dt.get();
                        final String previewTxt = PREVIEWBORDER_MSG + (newPreviewDataTable != null
                                ? ": " + newPreviewDataTable.getDataTableSpec().getName()
                                : "");
                        setPreviewTablePanelBorderTitle(previewTxt);
                        setNewPreviewDataTableInEDT(newPreviewDataTable);
                    } catch (final Throwable t) {
                        LOGGER.debug(t);
                        m_previewMsg.setText(messagePrefix.isEmpty() || (t.getMessage() == null) ? t.getMessage()
                            : messagePrefix + " " + t.getMessage());
                    }
                } finally {
                    // enable the refresh button again
                    m_previewUpdateButton.setEnabled(!m_noPreviewChecker.isSelected());
                }
            }
        };
        sw.execute();
    }

    /**
     *
     */
    void sheetNotFound() {
        String msg = "Could not load file";
        if (m_fileAccessError != null) {
            msg += ": " + m_fileAccessError;
        }
        m_previewMsg.setText(msg);
        clearTableViews();
        // enable the refresh button again
        m_previewUpdateButton.setEnabled(!m_noPreviewChecker.isSelected());
    }

    private XLSUserSettings createSettingsFromComponents() throws InvalidSettingsException {
        final XLSUserSettings s = new XLSUserSettings();
        String sheetName = (String)m_sheetName.getSelectedItem();
        if (FIRST_SHEET.equals(sheetName)) {
            sheetName = null;
        }
        s.setSheetName(sheetName);

        s.setSkipEmptyColumns(m_skipEmptyCols.isSelected());
        s.setSkipEmptyRows(m_skipEmptyRows.isSelected());
        s.setSkipHiddenColumns(m_skipHiddenColumns.isSelected());
        s.setReadAllData(m_readAllData.isSelected());

        s.setHasColHeaders(m_hasColHdr.isSelected());
        try {
            s.setColHdrRow(getPositiveNumberFromTextField(m_colHdrRow));
        } catch (final InvalidSettingsException ise) {
            if (m_hasColHdr.isSelected()) {
                throw new InvalidSettingsException("Column Header Row: " + ise.getMessage());
            }
            s.setColHdrRow(0);
        }
        s.setUniquifyRowIDs(m_uniquifyRowIDs.isSelected());
        s.setHasRowHeaders(m_hasRowIDs.isSelected());
        s.setIndexContinuous(m_indexContinuous.isSelected());
        s.setIndexSkipJumps(m_indexSkipJumps.isSelected());
        try {
            s.setRowHdrCol(getColumnNumberFromTextField(m_rowIDCol));
        } catch (final InvalidSettingsException ise) {
            if (m_hasRowIDs.isSelected()) {
                throw new InvalidSettingsException("Row Header Column Idx: " + ise.getMessage());
            }
            s.setRowHdrCol(0);
        }
        try {
            s.setFirstColumn(getColumnNumberFromTextField(m_firstCol));
        } catch (final InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("First Column: " + ise.getMessage());
            }
            s.setFirstColumn(0);
        }
        try {
            s.setLastColumn(getColumnNumberFromTextField(m_lastCol));
        } catch (final InvalidSettingsException ise) {
            // no last column specified
            s.setLastColumn(0);
        }
        try {
            s.setFirstRow(getPositiveNumberFromTextField(m_firstRow));
        } catch (final InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("First Row: " + ise.getMessage());
            }
            s.setFirstRow(0);
        }
        try {
            s.setLastRow(getPositiveNumberFromTextField(m_lastRow));
        } catch (final InvalidSettingsException ise) {
            // no last row set
            s.setLastRow(0);
        }

        // formula eval err handling
        s.setUseErrorPattern(m_formulaStringCell.isSelected());
        s.setErrorPattern(m_formulaErrPattern.getText());

        s.setReevaluateFormulae(m_reevaluateFormulae.isSelected());
        s.setTimeoutInSeconds(readTimeOutInSecondsFromSpinner());
        s.setNoPreview(m_noPreviewChecker.isSelected());
        return s;
    }

    /**
     * Creates an int from the specified text field. Throws a ISE if the entered value is empty, is not a number or zero
     * or negative.
     */
    private static int getPositiveNumberFromTextField(final JTextField t) throws InvalidSettingsException {
        final String input = t.getText();
        if ((input == null) || input.isEmpty()) {
            throw new InvalidSettingsException("please enter a number.");
        }
        int i;
        try {
            i = Integer.parseInt(input);
        } catch (final NumberFormatException nfe) {
            throw new InvalidSettingsException("not a valid integer number.");
        }
        if (i <= 0) {
            throw new InvalidSettingsException("number must be larger than zero.");
        }
        return i;
    }

    /**
     * Creates an int ({@code 1}-based) from the specified text field. It accepts numbers between 1 and 1024 (incl.) or
     * XLS column headers (starting at 'A', 'B', ... 'Z', 'AA', etc.) Throws a ISE if the entered value is not valid.
     */
    private static int getColumnNumberFromTextField(final JTextField t) throws InvalidSettingsException {
        final String input = t.getText();
        return POIUtils.oneBasedColumnNumberChecked(input);
    }

    private void transferSettingsIntoComponents(final XLSUserSettings s) {
        CheckUtils.checkState(
            (s.getHasRowHeaders() && !s.isIndexContinuous() && !s.isIndexSkipJumps())
                || (!s.getHasRowHeaders() && s.isIndexContinuous() && !s.isIndexSkipJumps())
                || (!s.getHasRowHeaders() && !s.isIndexContinuous() && s.isIndexSkipJumps()),
            "Exactly one of generate row ids or table contains row ids in column should be selected!");

        m_skipEmptyCols.setSelected(s.getSkipEmptyColumns());
        m_skipHiddenColumns.setSelected(s.getSkipHiddenColumns());
        m_skipEmptyRows.setSelected(s.getSkipEmptyRows());
        m_readAllData.setSelected(s.getReadAllData());

        // dialog shows numbers - internally we use indices
        m_hasColHdr.setSelected(s.getHasColHeaders());
        m_colHdrRow.setText(String.valueOf(s.getColHdrRow()));
        m_indexContinuous.setSelected(s.isIndexContinuous());
        m_indexSkipJumps.setSelected(s.isIndexSkipJumps());
        m_hasRowIDs.setSelected(s.getHasRowHeaders());
        m_uniquifyRowIDs.setSelected(s.getUniquifyRowIDs());

        int val;
        val = s.getRowHdrCol(); // getColLabel wants an index
        if (val >= 1) {
            m_rowIDCol.setText(POIUtils.oneBasedColumnNumber(val));
        } else {
            m_rowIDCol.setText("A");
        }
        val = s.getFirstColumn(); // getColLabel wants an index
        if (val >= 1) {
            m_firstCol.setText(POIUtils.oneBasedColumnNumber(val));
        } else {
            m_firstCol.setText("A");
        }
        val = s.getLastColumn();
        if (val >= 1) {
            m_lastCol.setText(POIUtils.oneBasedColumnNumber(val));
        } else {
            m_lastCol.setText("");
        }
        val = s.getFirstRow();
        if (val >= 1) {
            m_firstRow.setText("" + val);
        } else {
            m_firstRow.setText("1");
        }
        val = s.getLastRow();
        if (val >= 1) {
            m_lastRow.setText("" + val);
        } else {
            m_lastRow.setText("");
        }
        // formula error handling
        m_formulaStringCell.setSelected(s.getUseErrorPattern());
        m_formulaMissCell.setSelected(!s.getUseErrorPattern());
        m_formulaErrPattern.setText(s.getErrorPattern());
        m_formulaErrPattern.setEnabled(s.getUseErrorPattern());

        m_reevaluateFormulae.setSelected(s.isReevaluateFormulae());

        m_noPreviewChecker.setSelected(s.isNoPreview());

        m_timeout.setValue(s.getTimeoutInSeconds());

        // clear sheet names
        m_sheetName.setModel(new DefaultComboBoxModel<>());
        // set new sheet names
        updateSheetListAndSelect(s.getSheetName());
        // set the en/disable state
        checkBoxChanged();

    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        checkPreviousFuture();
        // we need at least a filename and sheet name
        final String file = m_fileChooserSettings.getPathOrURL();
        if ((file == null) || file.isEmpty()) {
            throw new InvalidSettingsException("Please select a file to read from.");
        }
        final String sheet = (String)m_sheetName.getSelectedItem();
        if (SCANNING.equals(sheet)) {
            throw new InvalidSettingsException("Please wait until the file scanning finishes and select a worksheet.");
        }
        if ((sheet == null) || sheet.isEmpty()) {
            throw new InvalidSettingsException("Please select a worksheet.");
        }

        final XLSUserSettings s = createSettingsFromComponents();
        final String errMsg = s.getStatus();
        if (errMsg != null) {
            throw new InvalidSettingsException(errMsg);
        }

        if (!m_previewMsg.getText().isEmpty() && !m_previewMsg.getText().startsWith(LOADING_INTERRUPTED)) {
            throw new InvalidSettingsException(m_previewMsg.getText());
        }
        s.save(settings);
        m_fileChooser.saveSettingsTo(settings);
        final DataTable preview = m_previewDataTable;
        if (!s.isNoPreview() /*&& !m_incomplete*/) {
            CheckUtils.checkSettingNotNull(preview, "No preview table created - reload the sheet");
            // if we have a preview table, store the DTS with the settings.
            // This is a hack around to avoid long configure times.
            // Causes the node's execute method to issue a bad warning, if the
            // file content changes between closing the dialog and execute()
            final String dtsId = SettingsIDBuilder.getID((SettingsModelFileChooser2)m_fileChooser.getModel(), s);
            settings.addString(XLS_CFG_ID_FOR_TABLESPEC,  dtsId);
            final Config subConf = settings.addConfig(XLS_CFG_TABLESPEC);
            preview.getDataTableSpec().save(subConf);
        }
    }

    /**
     * Sets {@link #m_isCurrentlyLoadingNodeSettings} and returns the previous value. It also enables/disables controls
     * according to the new state.
     *
     * @param newValue The new value to set
     * @return the previous value.
     */
    private boolean setCurrentlyLoadingNodeSettings(final boolean newValue) {
        final boolean oldValue = m_isCurrentlyLoadingNodeSettings;
        m_isCurrentlyLoadingNodeSettings = newValue;
        if (oldValue != newValue) {
            final boolean enableControls = !newValue;
            m_skipEmptyCols.setEnabled(enableControls);
            m_skipHiddenColumns.setEnabled(enableControls);
            m_skipEmptyRows.setEnabled(enableControls);
            m_reevaluateFormulae.setEnabled(enableControls);
            m_noPreviewChecker.setEnabled(enableControls);
            m_previewUpdateButton.setEnabled(enableControls);
            m_cancel.setEnabled(enableControls);
            onNoPreviewCheckerSelected();
        }
        return oldValue;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        m_fsConnection = FileSystemPortObjectSpec.getFileSystemConnection(specs, 0);

        setCurrentlyLoadingNodeSettings(true);
        clearTableViews();
        XLSUserSettings s;
        try {
            s = XLSUserSettings.load(settings);
        } catch (final InvalidSettingsException e) {
            s = new XLSUserSettings();
        }
        m_fileChooser.loadSettingsFrom(settings, specs);
        // Get the workbook when dialog is opened
        refreshWorkbook(((SettingsModelFileChooser2)m_fileChooser.getModel()).getPathOrURL());
        transferSettingsIntoComponents(s);
        if (FIRST_SHEET.equals(m_sheetName.getSelectedItem()) || (m_sheetName.getSelectedItem() == null)) {
            // now refresh preview tables
            updateFileTable();
        }
        setCurrentlyLoadingNodeSettings(false);
    }

    private void addFocusLostListener(final JTextField field) {
        field.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(final FocusEvent e) {
                invalidatePreviewTable();
            }
        });
    }

    private void refreshWorkbook(final String path) {
        if (path == null) {
            m_workbook = null;
            m_sheets.clear();
            m_workbookPath = null;
        } else if (!path.equals(m_workbookPath)) {
            m_sheets.clear();
            m_workbook = null;
            m_workbookPath = path;
            checkPreviousFuture();
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void onClose() {
        cleanup();
        super.onClose();
    }

    /**
     *
     */
    private void cleanup() {
        // Remove references to XLSTable, that holds a reference to the workbook
        clearTableViews();
        // Remove own reference to the workbook
        m_sheets.clear();
        m_workbook = null;
        m_workbookPath = null;
        checkPreviousFuture();
        // Now the garbage collector should be able to collect the workbook object
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void onCancel() {
        cleanup();
        super.onCancel();
    }

    /**
     * Special {@link TableContentViewTableHeader} to modify its tooltip.
     */
    @SuppressWarnings("serial")
    private final class TableContentViewTableHeaderWithTooltip extends TableContentViewTableHeader {

        private final Map<Integer, Integer> m_mapKNIMEToExcel;

        /**
         * @param contentView
         * @param cm
         */
        TableContentViewTableHeaderWithTooltip(final TableContentView contentView, final TableColumnModel cm,
            final Map<Integer, Integer> mapKNIMEToExcel) {
            super(contentView, cm);
            m_mapKNIMEToExcel = mapKNIMEToExcel;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public String getToolTipText(final MouseEvent event) {
            final int column = columnAtPoint(event.getPoint());
            if ((column >= 0) && m_mapKNIMEToExcel.containsKey(column)) {
                final int excelColumn = m_mapKNIMEToExcel.get(column) + 1;
                return POIUtils.oneBasedColumnNumber(excelColumn);
            }
            return super.getToolTipText(event);
        }
    }

    /**
     * @return
     */
    private String interruptedMessage() {
        final int nrRowsRead = m_readRows.intValue();
        if (nrRowsRead < 0) {
            String interruptSuffix;
            if (nrRowsRead == -1) {
                interruptSuffix = ", no preview is available";
            } else {
                interruptSuffix = ", preview is based on the first " + -(nrRowsRead + 1) + " values of the sheet";
            }
            return LOADING_INTERRUPTED + interruptSuffix;
        } else {
            return "";
        }
    }
}
