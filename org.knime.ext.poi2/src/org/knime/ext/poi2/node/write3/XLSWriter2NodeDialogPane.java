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
 *   Mar 16, 2007 (ohl): created
 */
package org.knime.ext.poi2.node.write3;

import java.awt.Dimension;
import java.awt.Window;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.knime.core.data.BooleanValue;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataValue;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.IntValue;
import org.knime.core.data.LongValue;
import org.knime.core.data.StringValue;
import org.knime.core.data.date.DateAndTimeValue;
import org.knime.core.data.image.png.PNGImageValue;
import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.FlowVariableModelButton;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.filter.NameFilterConfiguration;
import org.knime.core.node.util.filter.column.DataColumnSpecFilterConfiguration;
import org.knime.core.node.util.filter.column.DataColumnSpecFilterPanel;
import org.knime.core.node.util.filter.column.DataTypeColumnFilter;
import org.knime.core.node.workflow.FlowVariable;
import org.knime.core.node.workflow.FlowVariable.Type;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.defaultnodesettings.DialogComponentFileChooser2;
import org.knime.filehandling.core.defaultnodesettings.FileSystemChoice;
import org.knime.filehandling.core.defaultnodesettings.SettingsModelFileChooser2;
import org.knime.filehandling.core.port.FileSystemPortObjectSpec;

/**
 *
 * @author ohl, University of Konstanz
 */
final class XLSWriter2NodeDialogPane extends NodeDialogPane {

    @SuppressWarnings("unchecked")
    private static final Class<? extends DataValue>[] ACCEPTED_TYPES = new Class[]{StringValue.class, IntValue.class,
        LongValue.class, DoubleValue.class, PNGImageValue.class, BooleanValue.class, DateAndTimeValue.class};

    private final XLSNodeType m_type;

    private final DialogComponentFileChooser2 m_fileChooserComponent;

    private final JCheckBox m_writeColHdr = new JCheckBox();

    private final JCheckBox m_writeRowHdr = new JCheckBox();

    private final JTextField m_missValue = new JTextField();

    private final JCheckBox m_overwriteOK = new JCheckBox("Overwrite existing file");

    private final JCheckBox m_evaluateFormula = new JCheckBox("Evaluate formulas on write");

    private final JCheckBox m_fileMustExist = new JCheckBox("Abort if file does not exist");

    private final JCheckBox m_doNotOverwriteSheet = new JCheckBox("Abort if sheet already exists");

    private final JCheckBox m_openFile = new JCheckBox("Open file after execution");

    private final JCheckBox m_autosize = new JCheckBox("Autosize columns");

    private final JTextField m_sheetname = new JTextField();

    private final ButtonGroup m_pageFormat = new ButtonGroup();

    private final JRadioButton m_portrait = new JRadioButton("Portrait");

    private final JRadioButton m_landscape = new JRadioButton("Landscape");

    private final JComboBox<String> m_paperSize = new JComboBox<>(PaperSize.getNames());

    private final FlowVariableModelButton m_sheetnameFVM =
        new FlowVariableModelButton(createFlowVariableModel("sheetname", FlowVariable.Type.STRING));

    private final DataColumnSpecFilterPanel m_filter = new DataColumnSpecFilterPanel();

    private final JCheckBox m_writeMissingValue = new JCheckBox("For missing values write:");

    private boolean m_xlsIsConfirmed = false;

    private Optional<FSConnection> m_fs;

    /**
     * Creates a new dialog for the XLS writer node.
     *
     *
     * @param type Of what type is this node?
     */
    XLSWriter2NodeDialogPane(final XLSNodeType type) {
        m_type = type;
        final JPanel tab = new JPanel();
        tab.setLayout(new BoxLayout(tab, BoxLayout.Y_AXIS));

        // enable xlsm support for file chooser if node is appender
        final String[] validExtensions =
            m_type == XLSNodeType.APPENDER ? new String[]{"xlsx", "xls", "xlsm"} : new String[]{"xlsx", "xls"};

        final SettingsModelFileChooser2 fileChooser =
            XLSWriter2NodeModel.createSettingsModelFileChooser(validExtensions);
        final FlowVariableModel fvm = createFlowVariableModel(
            new String[]{fileChooser.getConfigName(), SettingsModelFileChooser2.PATH_OR_URL_KEY}, Type.STRING);

        m_fileChooserComponent = new DialogComponentFileChooser2(1, fileChooser, "XLSWRITER", JFileChooser.SAVE_DIALOG,
            JFileChooser.FILES_ONLY, fvm);
        m_fileChooserComponent.setForceExtensionOnSave(".xlsx");

        final JPanel filePanel = new JPanel();
        filePanel.setLayout(new BoxLayout(filePanel, BoxLayout.X_AXIS));
        filePanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Output location:"));
        filePanel.setMaximumSize(
            new Dimension(Integer.MAX_VALUE, m_fileChooserComponent.getComponentPanel().getPreferredSize().height));
        filePanel.add(m_fileChooserComponent.getComponentPanel());
        filePanel.add(Box.createHorizontalGlue());

        tab.add(filePanel);

        if (m_type == XLSNodeType.WRITER) {
            tab.add(createFileOverwriteBox());
        }
        if (m_type == XLSNodeType.APPENDER) {
            tab.add(createFileMustExistBox());
            tab.add(createDoNotOverwriteSheetBox());
        }
        tab.add(createOpenFileBox());
        tab.add(createSheetnameBox());
        tab.add(Box.createVerticalStrut(5));
        tab.add(createHeadersBox());
        if (m_type == XLSNodeType.APPENDER) {
            tab.add(createFormulaBox());
        }
        tab.add(createMissingBox());
        tab.add(createLayoutBox());
        tab.add(m_filter);

        addTab("Settings", tab);

        m_writeMissingValue.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(final ChangeEvent e) {
                m_missValue.setEnabled(m_writeMissingValue.isSelected());
            }
        });
    }

    private JComponent createFileOverwriteBox() {
        final Box overwiteBox = Box.createHorizontalBox();
        overwiteBox.add(m_overwriteOK);
        overwiteBox.add(Box.createHorizontalGlue());
        overwiteBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_overwriteOK.getMaximumSize().height));
        return overwiteBox;
    }

    private JComponent createFormulaBox() {
        final Box formulaBox = Box.createHorizontalBox();
        formulaBox.add(m_evaluateFormula);
        formulaBox.add(Box.createHorizontalGlue());
        formulaBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_evaluateFormula.getMaximumSize().height));

        final JPanel result = new JPanel();
        result.setLayout(new BoxLayout(result, BoxLayout.Y_AXIS));
        result.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Formulas"));
        result.add(formulaBox);

        return result;
    }

    private JComponent createFileMustExistBox() {
        final Box fileMustExistBox = Box.createHorizontalBox();
        fileMustExistBox.add(m_fileMustExist);
        fileMustExistBox.add(Box.createHorizontalGlue());
        fileMustExistBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_fileMustExist.getMaximumSize().height));
        return fileMustExistBox;
    }

    private JComponent createDoNotOverwriteSheetBox() {
        final Box doNotOverwriteSheetBox = Box.createHorizontalBox();
        doNotOverwriteSheetBox.add(m_doNotOverwriteSheet);
        doNotOverwriteSheetBox.add(Box.createHorizontalGlue());
        doNotOverwriteSheetBox
            .setMaximumSize(new Dimension(Integer.MAX_VALUE, m_doNotOverwriteSheet.getMaximumSize().height));
        return doNotOverwriteSheetBox;
    }

    private JComponent createOpenFileBox() {
        final Box openfileBox = Box.createHorizontalBox();
        openfileBox.add(m_openFile);
        openfileBox.add(Box.createHorizontalGlue());
        openfileBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_openFile.getMaximumSize().height));
        return openfileBox;
    }

    private JComponent createSheetnameBox() {
        final Box sheetnameBox = Box.createHorizontalBox();
        sheetnameBox.setBorder(new TitledBorder(new EtchedBorder(), "Sheet name"));
        m_sheetnameFVM.getFlowVariableModel().addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(final ChangeEvent arg0) {
                m_sheetname.setEnabled(!m_sheetnameFVM.getFlowVariableModel().isVariableReplacementEnabled());
            }
        });
        m_sheetname.setPreferredSize(new Dimension(410, 25));
        m_sheetname.setMaximumSize(new Dimension(Integer.MAX_VALUE, 25));
        sheetnameBox.add(new JLabel("Name of the sheet:"));
        sheetnameBox.add(Box.createHorizontalStrut(5));
        sheetnameBox.add(m_sheetname);
        sheetnameBox.add(Box.createHorizontalStrut(5));
        sheetnameBox.add(m_sheetnameFVM);
        sheetnameBox.add(Box.createHorizontalGlue());
        sheetnameBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_sheetname.getMaximumSize().height));
        return sheetnameBox;
    }

    private JPanel createHeadersBox() {

        m_writeColHdr.setText("add column headers");
        m_writeColHdr.setToolTipText("Column names are stored in the first row" + " of the datasheet");
        m_writeRowHdr.setText("add row ids");
        m_writeRowHdr.setToolTipText("Row IDs are stored in the first column" + " of the datasheet");

        final Box colHdrBox = Box.createHorizontalBox();
        colHdrBox.add(m_writeColHdr);
        colHdrBox.add(Box.createHorizontalGlue());
        colHdrBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_writeColHdr.getMaximumSize().height));

        final Box rowHdrBox = Box.createHorizontalBox();
        rowHdrBox.add(m_writeRowHdr);
        rowHdrBox.add(Box.createHorizontalGlue());
        rowHdrBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_writeRowHdr.getMaximumSize().height));

        final JPanel result = new JPanel();
        result.setLayout(new BoxLayout(result, BoxLayout.Y_AXIS));
        result.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Add names and IDs"));
        result.add(colHdrBox);
        result.add(rowHdrBox);

        return result;

    }

    private JPanel createMissingBox() {

        m_missValue.setToolTipText("This text will be set for missing values." + "If unsure, leave empty");
        m_missValue.setPreferredSize(new Dimension(400, 25));
        m_missValue.setMaximumSize(new Dimension(Integer.MAX_VALUE, 25));
        final Box missBox = Box.createHorizontalBox();
        missBox.add(m_writeMissingValue);
        missBox.add(Box.createHorizontalStrut(5));
        missBox.add(m_missValue);
        missBox.add(Box.createHorizontalGlue());
        missBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_missValue.getMaximumSize().height));

        final JPanel result = new JPanel();
        result.setLayout(new BoxLayout(result, BoxLayout.Y_AXIS));
        result.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Missing value pattern"));
        result.add(missBox);
        return result;

    }

    private JPanel createLayoutBox() {
        m_portrait.setActionCommand(m_portrait.getText());
        m_portrait.setSelected(true);
        m_landscape.setActionCommand(m_landscape.getText());
        m_pageFormat.add(m_portrait);
        m_pageFormat.add(m_landscape);
        m_paperSize.setMaximumSize(m_paperSize.getPreferredSize());

        final Box sizeBox = Box.createHorizontalBox();
        sizeBox.add(m_autosize);
        sizeBox.add(Box.createHorizontalGlue());
        sizeBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_autosize.getMaximumSize().height));

        final Box formatBox = Box.createHorizontalBox();
        formatBox.add(m_portrait);
        formatBox.add(m_landscape);
        formatBox.add(Box.createHorizontalStrut(15));
        formatBox.add(m_paperSize);
        formatBox.add(Box.createHorizontalGlue());
        formatBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, m_portrait.getMaximumSize().height));

        final JPanel result = new JPanel();
        result.setLayout(new BoxLayout(result, BoxLayout.Y_AXIS));
        result.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Layout"));
        result.add(sizeBox);
        result.add(Box.createVerticalStrut(5));
        result.add(formatBox);
        result.add(Box.createVerticalStrut(5));
        return result;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        final DataTableSpec tableSpec = (DataTableSpec)specs[0];
        m_fs = FileSystemPortObjectSpec.getFileSystemConnection(specs, 1);

        addChangeListenerToFileChooserSettings();

        XLSWriter2Settings newVals;
        try {
            newVals = new XLSWriter2Settings(settings);
        } catch (final InvalidSettingsException ise) {
            // keep the defaults.
            newVals = new XLSWriter2Settings(tableSpec);
        }
        //Necessary to update the file dialog component
        m_fileChooserComponent.loadSettingsFrom(settings, specs);
        m_missValue.setText(newVals.getMissingPattern());
        m_writeColHdr.setSelected(newVals.writeColHeader());
        m_writeRowHdr.setSelected(newVals.writeRowID());
        m_evaluateFormula.setSelected(newVals.evaluateFormula());
        m_overwriteOK.setSelected(newVals.getOverwriteOK());
        m_openFile.setSelected(newVals.getOpenFile());
        m_sheetname.setText(newVals.getSheetname());
        m_sheetname.setEnabled(!m_sheetnameFVM.getFlowVariableModel().isVariableReplacementEnabled());
        m_fileMustExist.setSelected(newVals.getFileMustExist());
        m_doNotOverwriteSheet.setSelected(newVals.getDoNotOverwriteSheet());
        m_autosize.setSelected(newVals.getAutosize());
        if (newVals.getLandscape()) {
            m_landscape.setSelected(true);
        } else {
            m_portrait.setSelected(true);
        }
        m_paperSize.setSelectedItem(PaperSize.getName(newVals.getPaperSize()));
        m_writeMissingValue.setSelected(newVals.getWriteMissingValues());
        m_missValue.setEnabled(newVals.getWriteMissingValues());
        final DataColumnSpecFilterConfiguration config = createColFilterConf();
        config.loadConfigurationInDialog(settings, tableSpec);
        m_filter.loadConfiguration(config, tableSpec);
    }

    private void addChangeListenerToFileChooserSettings() {
        final SettingsModelFileChooser2 model = (SettingsModelFileChooser2)m_fileChooserComponent.getModel();
        model.addChangeListener(
            event -> m_openFile.setEnabled(model.getFileSystemChoice() == FileSystemChoice.getLocalFsChoice()));
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {

        final String filename = ((SettingsModelFileChooser2)m_fileChooserComponent.getModel()).getPathOrURL();

        final String fileExtension = FilenameUtils.getExtension(filename).toLowerCase();
        if (fileExtension.equals("xls")) {
            if (!m_xlsIsConfirmed) {
                final Window windowAncestor =
                    SwingUtilities.getWindowAncestor(m_fileChooserComponent.getComponentPanel());
                final int result = JOptionPane.showOptionDialog(windowAncestor,
                    "<html>You chose the the old <i>xls</i> "
                        + "Excel file format, which was used as default up until Excel 2003. This file format has "
                        + "certain limitations. <br/>"
                        + "Do you still want to write <i>xls</i> or use the newer <i>xlsx</i> format? "
                        + "For details see the node description.</html>",
                    ".xls or .xlsx?", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null,
                    new Object[]{"Write 'XLSX'", "Write 'XLS'"}, "Write 'XLSX'");
                switch (result) {
                    case JOptionPane.YES_OPTION:
                        ((SettingsModelFileChooser2)m_fileChooserComponent.getModel())
                            .setPathOrURL((FilenameUtils.removeExtension(filename) + ".xlsx"));
                        break;
                    case JOptionPane.NO_OPTION:
                        m_xlsIsConfirmed = true;
                        break;
                    default:
                        //Nothing to do
                }
            }
        } else if (!isValidFileExtension(fileExtension) && !FilenameUtils.getBaseName(filename).trim().isEmpty()) {
                //Not empty, but no valid extension, add extension
                ((SettingsModelFileChooser2)m_fileChooserComponent.getModel()).setPathOrURL(filename + ".xlsx");
        }

        m_fileChooserComponent.saveSettingsTo(settings);
        final XLSWriter2Settings vals = new XLSWriter2Settings();
        vals.setMissingPattern(m_missValue.getText());
        vals.setWriteColHeader(m_writeColHdr.isSelected());
        vals.setWriteRowID(m_writeRowHdr.isSelected());
        vals.setOverwriteOK(m_overwriteOK.isSelected());
        vals.setOpenFile(m_openFile.isSelected());
        vals.setSheetname(m_sheetname.getText());
        vals.setFileMustExist(m_fileMustExist.isSelected());
        vals.setDoNotOverwriteSheet(m_doNotOverwriteSheet.isSelected());
        vals.setAutosize(m_autosize.isSelected());
        vals.setLandscape(m_landscape.isSelected());
        vals.setPaperSize(PaperSize.getValue((String)m_paperSize.getSelectedItem()));
        vals.setWriteMissingValues(m_writeMissingValue.isSelected());
        vals.setEvaluateFormula(m_evaluateFormula.isSelected());
        vals.saveSettingsTo(settings);
        final DataColumnSpecFilterConfiguration config = createColFilterConf();
        m_filter.saveConfiguration(config);
        config.saveConfiguration(settings);
    }

    private boolean isValidFileExtension(final String fileExtension) {
        boolean result = false;
        if (m_type.equals(XLSNodeType.WRITER)) {
            result = fileExtension.equals("xlsx");
        } else if (m_type.equals(XLSNodeType.APPENDER)) {
            result = fileExtension.equals("xlsx") || fileExtension.equals("xlsm");
        }
        return result;
    }

    /**
     * @return creates and returns configuration instance for column filter panel.
     */
    static DataColumnSpecFilterConfiguration createColFilterConf() {
        return new DataColumnSpecFilterConfiguration("xlswriter2", new DataTypeColumnFilter(ACCEPTED_TYPES),
            NameFilterConfiguration.FILTER_BY_NAMEPATTERN | DataColumnSpecFilterConfiguration.FILTER_BY_DATATYPE);
    }

    private static class PaperSize {

        private static final Map<String, Short> SIZES = new HashMap<String, Short>();

        static {
            SIZES.put("A4 - 210x297 mm", PrintSetup.A4_PAPERSIZE);
            SIZES.put("A5 - 148x210 mm", PrintSetup.A5_PAPERSIZE);
            SIZES.put("US Envelope #10 4 1/8 x 9 1/2", PrintSetup.ENVELOPE_10_PAPERSIZE);
            SIZES.put("Envelope C5 162x229 mm", PrintSetup.ENVELOPE_CS_PAPERSIZE);
            SIZES.put("Envelope DL 110x220 mm", PrintSetup.ENVELOPE_DL_PAPERSIZE);
            SIZES.put("Envelope Monarch 98.4×190.5 mm", PrintSetup.ENVELOPE_MONARCH_PAPERSIZE);
            SIZES.put("US Executive 7 1/4 x 10 1/2 in", PrintSetup.EXECUTIVE_PAPERSIZE);
            SIZES.put("US Legal 8 1/2 x 14 in", PrintSetup.LEGAL_PAPERSIZE);
            SIZES.put("US Letter 8 1/2 x 11 in", PrintSetup.LETTER_PAPERSIZE);
        }

        static String[] getNames() {
            return SIZES.keySet().toArray(new String[SIZES.size()]);
        }

        static short getValue(final String name) {
            return SIZES.get(name).shortValue();
        }

        static String getName(final short value) {
            for (final Entry<String, Short> entry : SIZES.entrySet()) {
                if (entry.getValue().shortValue() == value) {
                    return entry.getKey();
                }
            }
            return null;
        }

    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void onOpen() {
        super.onOpen();
        m_xlsIsConfirmed = false;
    }
}
