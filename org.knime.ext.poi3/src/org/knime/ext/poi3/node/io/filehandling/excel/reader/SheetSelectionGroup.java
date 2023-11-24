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
 *   24 Nov 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionListener;
import java.util.EventObject;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.SpinnerNumberModel;
import javax.swing.event.ChangeListener;

import org.knime.ext.poi3.node.io.filehandling.excel.DialogUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.util.GBCBuilder;

/**
 * Bundle of controls for selecting the sheet that is read by the dialog.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class SheetSelectionGroup {

    // Code extracted from ExcelTableReaderNodeDialog

    private final JRadioButton m_radioButtonFirstSheetWithData = new JRadioButton(SheetSelection.FIRST.getText(), true);

    private final JRadioButton m_radioButtonSheetByName = new JRadioButton(SheetSelection.NAME.getText());

    private final JRadioButton m_radioButtonSheetByIndex = new JRadioButton(SheetSelection.INDEX.getText());

    private final ButtonGroup m_sheetSelectionButtonGroup = new ButtonGroup();

    private final JLabel m_firstSheetWithDataLabel = new JLabel();

    private JComboBox<String> m_sheetNameSelection = new JComboBox<>();

    private JSpinner m_sheetIndexSelection = new JSpinner(new SpinnerNumberModel(0, 0, Integer.MAX_VALUE, 1));

    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_sheetIdxNoteLabel = new JLabel("(Position start with 0.) ");

    private final Supplier<FilterMode> m_currentFilterMode;

    SheetSelectionGroup(final Supplier<FilterMode> currentFilterMode) {
        m_currentFilterMode = currentFilterMode;
        DialogUtil.italicizeText(m_firstSheetWithDataLabel);
        DialogUtil.italicizeText(m_sheetIdxNoteLabel);
        m_sheetSelectionButtonGroup.add(m_radioButtonFirstSheetWithData);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByName);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByIndex);
    }

    public void registerDialogChangeListeners() {
        m_radioButtonSheetByName.addChangeListener(l -> m_sheetNameSelection
            .setEnabled(m_radioButtonSheetByName.isEnabled() && m_radioButtonSheetByName.isSelected()));
        m_radioButtonSheetByIndex.addChangeListener(l -> m_sheetIndexSelection
            .setEnabled(m_radioButtonSheetByIndex.isEnabled() && m_radioButtonSheetByIndex.isSelected()));
    }

    /**
     * Resets the current selection.
     */
    public void resetSelection() {
        // TODO memorize and reconsile later
        m_radioButtonFirstSheetWithData.setSelected(true);
        m_radioButtonSheetByName.setEnabled(false);
        m_radioButtonSheetByIndex.setEnabled(false);
        m_firstSheetWithDataLabel.setText(null);
    }

    /**
     * @param sheetNames
     */
    public void setSheetNameList(final Map<String, Boolean> sheetNames) {

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
        if (m_currentFilterMode.get() == FilterMode.FILE) {
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
    }

    JPanel createFileAndSheetSelectSheetPanel(final String title) {
        final JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), title));
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

    public void registerConfigRelevantForFileContentChangedListener(final Consumer<EventObject> eventConsumer) {
        final ChangeListener cl = eventConsumer::accept;
        final ActionListener al = eventConsumer::accept;
        m_sheetIndexSelection.addChangeListener(cl);
        m_radioButtonFirstSheetWithData.addActionListener(al);
        m_radioButtonSheetByName.addActionListener(al);
        m_radioButtonSheetByIndex.addActionListener(al);
        m_sheetNameSelection.addActionListener(al);
    }

    public void saveToExcelConfig(final ExcelTableReaderConfig excelConfig) {
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

    public void loadExcelSettings(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getSheetSelection()) {
            case NAME -> m_radioButtonSheetByName.setSelected(true);
            case INDEX -> m_radioButtonSheetByIndex.setSelected(true);
            case FIRST -> m_radioButtonFirstSheetWithData.setSelected(true);
        }
        m_sheetNameSelection.addItem(excelConfig.getSheetName());
        m_sheetNameSelection.setSelectedItem(excelConfig.getSheetName());
        m_sheetIndexSelection.setValue(excelConfig.getSheetIdx());
    }

    public void updateFileContentPreviewSettings(final ExcelTableReaderConfig excelConfig) {
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

}
