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

import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.Arrays;
import java.util.EventObject;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.SpinnerNumberModel;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.plaf.basic.BasicComboBoxRenderer;

import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.DialogUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.util.GBCBuilder;

/**
 * Dialog controls for selecting the sheet that is read by the node.
 * The component handles marking missing sheets in the dropdown.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class SheetSelectionComponent implements ChangeListener {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(SheetSelectionComponent.class);

    private record SheetName(String label, boolean isMissing) {

        SheetName {
            if (label.isBlank()) {
                // catch programming errors where the UI ends up looking broken
                // (Excel does not allow blank sheet names anyway)
                throw new IllegalArgumentException("Sheet name may not be blank");
            }
        }
    }

    @SuppressWarnings("serial")
    private static final class SheetNamesComboBoxModel extends DefaultComboBoxModel<SheetName> {

        /**
         * Get the {@link SheetName} if it is already in the list, otherwise add it as non-missing.
         * @param sheetName name of the sheet
         * @return existing or newly added sheet
         */
        private SheetName getOrAddSheetName(final String sheetName) {
            final var index = indexOf(sheetName);
            if (index >= 0) {
                return getElementAt(index);
            }
            final var item = new SheetName(sheetName, false);
            addElement(item);
            return item;
        }

        private int indexOf(final String sheetName) {
            final var total = getSize();
            if (total <= 0) {
                return -1;
            }
            for (var i = 0; i < total; i++) {
                final var item = getElementAt(i);
                if (sheetName.equals(item.label())) {
                    return i;
                }
            }
            return -1;
        }
    }

    private final ButtonGroup m_sheetSelectionButtonGroup = new ButtonGroup();

    private final JRadioButton m_radioButtonFirstSheetWithData = new JRadioButton(SheetSelection.FIRST.getText(), true);
    private final JLabel m_firstSheetWithDataLabel = new JLabel();

    private final JRadioButton m_radioButtonSheetByName = new JRadioButton(SheetSelection.NAME.getText());
    private final SheetNamesComboBoxModel m_sheetNameSelectionModel = new SheetNamesComboBoxModel();
    private final JComboBox<SheetName> m_sheetNameSelection = new JComboBox<>(m_sheetNameSelectionModel);

    private final JRadioButton m_radioButtonSheetByIndex = new JRadioButton(SheetSelection.INDEX.getText());
    private final JSpinner m_sheetIndexSelection = new JSpinner(new SpinnerNumberModel(0, 0, Integer.MAX_VALUE, 1));
    // text has trailing space to not be cut on Windows because of italic font
    private final JLabel m_sheetIdxNoteLabel = new JLabel("(Position starts with 0.) ");

    private final Supplier<FilterMode> m_currentFilterMode;

    private State m_currentState;

    private enum State {
        /**
         * State for a new selection group, before settings are loaded into it.
         * Can only transition to {@link #READY} via loading settings.
         */
        NEW,
        /**
         * State when the group is configured after loading settings or receiving new sheet names.
         * Being in this state does not necessarily mean the group is consistent,
         * i.e. some sheet names in the list can be missing in the workbook.
         */
        READY,
        /**
         * State while waiting for new sheet names.
         */
        WAITING;
    }

    private void transitionTo(final State[] validCurrentStates, final State targetState) {
        if (LOGGER.isDebugEnabled()) {
            try (final var sw = new StringWriter(); final var pw = new PrintWriter(sw)) {
                new Exception().printStackTrace(pw);
                if (Arrays.stream(validCurrentStates).noneMatch(m_currentState::equals)) {
                    final var exp = Arrays.stream(validCurrentStates)
                            .map(State::name)
                            .collect(Collectors.joining(","));
                    LOGGER.debug("Expected to be any of [%s] instead of %s%n%s"
                        .formatted(exp, m_currentState, sw.toString()));
                } else {
                    LOGGER.debug("Normal transition from %s to %s%n%s"
                        .formatted(m_currentState, targetState, sw.toString()));
                }
            } catch (final IOException e) {
                // should not happen
                LOGGER.warn(e);
            }
        }
        m_currentState = targetState;
    }

    SheetSelectionComponent(final Supplier<FilterMode> currentFilterMode) {
        m_currentFilterMode = currentFilterMode;
        DialogUtil.italicizeText(m_firstSheetWithDataLabel);
        DialogUtil.italicizeText(m_sheetIdxNoteLabel);
        m_sheetSelectionButtonGroup.add(m_radioButtonFirstSheetWithData);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByName);
        m_sheetSelectionButtonGroup.add(m_radioButtonSheetByIndex);

        // custom renderer for "missing" SheetName list items
        final var sheetNameComboBoxRenderer = new SheetNameRenderer();
        m_sheetNameSelection.setRenderer(sheetNameComboBoxRenderer);
        // when dropdown is not open, ListItemCellRenderer's foreground color is not shown, so we have to set it
        // on the combobox level via this handler
        SheetNameRenderer.registerComboBoxSelectedMissingHandler(m_sheetNameSelection);

        m_radioButtonSheetByName.addChangeListener(l -> m_sheetNameSelection
            .setEnabled(m_radioButtonSheetByName.isEnabled() && m_radioButtonSheetByName.isSelected()));
        m_radioButtonSheetByIndex.addChangeListener(l -> m_sheetIndexSelection
            .setEnabled(m_radioButtonSheetByIndex.isEnabled() && m_radioButtonSheetByIndex.isSelected()));
        m_currentState = State.NEW;
    }

    @Override
    public void stateChanged(final ChangeEvent e) {
        LOGGER.debug("Sheet selection notified: %s".formatted(e));
    }

    /**
     * Puts the dialog components into the "updating" state, i.e. resets the current selection,
     * remembering it for later. The state is reconsiled later at the end of {@link #updateSheetSelection(Map)}.
     */
    public void prepareUpdateSheetSelection() {
        transitionTo(new State[] { State.READY, State.WAITING }, State.WAITING );
        m_firstSheetWithDataLabel.setText(null);
        // Note: we intentionally never disable "first sheet with data" option
        m_radioButtonSheetByName.setEnabled(false);
        m_radioButtonSheetByIndex.setEnabled(false);
    }

    /**
     * Set the sheet name list and update the dialog elements accordingly.
     *
     * @param sheetNames list of sheet names with flag indicating first non-empty sheet
     */
    public void updateSheetSelection(final Map<String, Boolean> sheetNames) {
        transitionTo(new State[] { State.WAITING }, State.READY);
        m_radioButtonFirstSheetWithData.setEnabled(true);
        // only enabled when the dropdown contains something to select
        m_radioButtonSheetByName.setEnabled(updateSheetNameDropdown(sheetNames));
        // there's no sub component to interact with, only a label shown, so nothing to re-enable
        updateFirstSheetWithDataText(sheetNames);
        // only enable when we actually have something to select by index
        updateSheetIndexSpinner(sheetNames.size());
        m_radioButtonSheetByIndex.setEnabled(!sheetNames.isEmpty());
    }

    /**
     * @param sheetNames list of sheet names with flag indicating first non-empty sheet
     */
    private void updateFirstSheetWithDataText(final Map<String, Boolean> sheetNames) {
        final String text;
        if (m_currentFilterMode.get() == FilterMode.FILE) {
            // text has trailing space to not be cut on Windows because of italic font
            text = ExcelUtils.getFirstSheetWithDataOrFirstIfAllEmpty(sheetNames)
                    .map(s -> String.format("(%s) ", s))
                    .orElse("<No sheet available> ");
        } else {
            text = "";
        }
        m_firstSheetWithDataLabel.setText(text);
        // set the preferred height to be correctly aligned with the radio button
        m_firstSheetWithDataLabel.setPreferredSize(new Dimension((int)new JLabel(text).getPreferredSize().getWidth(),
            (int)m_radioButtonFirstSheetWithData.getPreferredSize().getHeight()));

    }

    private void updateSheetIndexSpinner(final int currentNumberOfSheets) {
        final int selectedIndex = (int)m_sheetIndexSelection.getValue();
        final int newIndex = selectedIndex > (currentNumberOfSheets - 1) ? 0 : selectedIndex;
        m_sheetIndexSelection.setModel(new SpinnerNumberModel(
            newIndex, 0, currentNumberOfSheets - 1, 1));
    }

    /**
     * Update the dropdown with the given sheet names.
     * @param sheetNames sheet names and flag indicating first non-empty sheet
     * @return {@code true} if the combobox model is non-empty after the update, {@code false} otherwise
     */
    private boolean updateSheetNameDropdown(final Map<String, Boolean> sheetNames) {
        // both empty -> NOOP
        boolean inputIsEmpty = sheetNames.isEmpty();
        boolean selectionModelIsEmpty = m_sheetNameSelectionModel.getSize() == 0;
        if (inputIsEmpty && selectionModelIsEmpty) {
            // nothing to do, not even to mark something as missing
            return false;
        }
        // empty input, but present combobox values -> CLEAR selection and put any selected item as missing
        if (inputIsEmpty) {
            final var selectedItem = (SheetName)m_sheetNameSelectionModel.getSelectedItem();
            m_sheetNameSelectionModel.removeAllElements();
            if (selectedItem != null) {
                // MISSING selected
                m_sheetNameSelectionModel.addElement(new SheetName(selectedItem.label(), true));
                m_sheetNameSelection.setSelectedIndex(0);
            }
            return m_sheetNameSelectionModel.getSize() != 0;
        }
        // non-empty input from here on out

        final var sheetNamesList = sheetNames.keySet().stream().map(s -> new SheetName(s, false)).toList();
        if (selectionModelIsEmpty) {
            m_sheetNameSelectionModel.addAll(sheetNamesList);
            if (!sheetNamesList.isEmpty()) {
                m_sheetNameSelection.setSelectedIndex(0);
            }
            return m_sheetNameSelectionModel.getSize() != 0;
        }
        // combobox is non-empty from here on out

        final var selectedItem = (SheetName)m_sheetNameSelectionModel.getSelectedItem();
        m_sheetNameSelectionModel.removeAllElements();
        // figure out if we need to prepend the selected item as missing or if it's still in the list
        final var selectedMissing = selectedItem != null && sheetNamesList.stream().map(SheetName::label)
                .noneMatch(l -> l.equals(selectedItem.label()));
        if (selectedItem != null && selectedMissing) {
            // something selected, but not in the new list
            // MISSING selected
            m_sheetNameSelectionModel.addElement(new SheetName(selectedItem.label, true));
            m_sheetNameSelection.setSelectedIndex(0);
        }

        m_sheetNameSelectionModel.addAll(sheetNamesList);
        if (selectedItem != null && !selectedMissing) {
            // something was selected, and it is still in the list
            final var index = m_sheetNameSelectionModel.indexOf(selectedItem.label());
            m_sheetNameSelection.setSelectedIndex(index);
        }
        return m_sheetNameSelectionModel.getSize() != 0;
    }

    JPanel createFileAndSheetSelectSheetPanel(final String title) {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), title));
        final var insets = new Insets(5, 5, 0, 0);
        final var gbcBuilder = new GBCBuilder(insets).resetPos().anchorFirstLineStart();

        panel.add(m_radioButtonFirstSheetWithData, gbcBuilder.build());
        panel.add(m_firstSheetWithDataLabel, gbcBuilder.incX().build());
        panel.add(m_radioButtonSheetByName, gbcBuilder.resetX().incY().setWeightX(0).build());
        panel.add(m_sheetNameSelection, gbcBuilder.incX().build());
        panel.add(m_radioButtonSheetByIndex, gbcBuilder.incY().resetX().build());
        final var smallDouble = 1E-10; // spinner should only fill up space of the combo box, rest is for label
        panel.add(m_sheetIndexSelection, gbcBuilder.incX().setWeightX(smallDouble).fillHorizontal().build());
        m_sheetIdxNoteLabel.setPreferredSize(new Dimension((int)m_sheetIdxNoteLabel.getPreferredSize().getWidth(),
            (int)m_sheetIndexSelection.getPreferredSize().getHeight()));
        panel.add(m_sheetIdxNoteLabel, gbcBuilder.incX().setWeightX(1 - smallDouble).build());
        return panel;
    }

    /**
     * Registers the given listener with components emitting changes relevant for file content.
     *
     * @param eventConsumer listener to register with components
     */
    public void registerConfigRelevantForFileContentChangedListener(final Consumer<EventObject> eventConsumer) {
        final ChangeListener cl = eventConsumer::accept;
        final ActionListener al = eventConsumer::accept;
        m_sheetIndexSelection.addChangeListener(cl);
        m_radioButtonFirstSheetWithData.addActionListener(al);
        m_radioButtonSheetByName.addActionListener(al);
        m_radioButtonSheetByIndex.addActionListener(al);
        m_sheetNameSelection.addActionListener(al);
    }

    /**
     * Saves the current dialog state to the given configuration object.
     *
     * @param excelConfig configuration object to save to
     */
    public void saveToExcelConfig(final ExcelTableReaderConfig excelConfig) {
        if (m_radioButtonSheetByName.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.NAME);
        } else if (m_radioButtonSheetByIndex.isSelected()) {
            excelConfig.setSheetSelection(SheetSelection.INDEX);
        } else {
            excelConfig.setSheetSelection(SheetSelection.FIRST);
        }
        final var sheetName = (SheetName)m_sheetNameSelection.getSelectedItem();
        excelConfig.setSheetName(sheetName != null ? sheetName.label() : null);
        excelConfig.setSheetIdx((int)m_sheetIndexSelection.getValue());
    }

    /**
     * Loads the new dialog state from the given configuration object.
     *
     * @param excelConfig configuration object to load from
     */
    public void loadExcelSettings(final ExcelTableReaderConfig excelConfig) {
        switch (excelConfig.getSheetSelection()) {
            case NAME -> m_radioButtonSheetByName.setSelected(true);
            case INDEX -> m_radioButtonSheetByIndex.setSelected(true);
            case FIRST -> m_radioButtonFirstSheetWithData.setSelected(true);
        }
        final var sheetName = excelConfig.getSheetName();
        if (sheetName != null && !sheetName.isEmpty()) {
            final var item = m_sheetNameSelectionModel.getOrAddSheetName(sheetName);
            m_sheetNameSelection.setSelectedItem(item);
        }
        m_sheetIndexSelection.setValue(excelConfig.getSheetIdx());
        transitionTo(new State[] { State.NEW, State.READY }, State.READY);
    }

    /**
     * ComboBox renderer for {@link SheetName}s that handles coloring "missing" list items.
     *
     * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
     */
    @SuppressWarnings("serial")
    private static final class SheetNameRenderer extends BasicComboBoxRenderer {
        /**
         * Colors a "missing" {@link SheetName} red and prepends the "missing" indicator text.
         */
        @Override
        public java.awt.Component getListCellRendererComponent(final javax.swing.JList<?> list, final Object value,
            final int index, final boolean isSelected, final boolean cellHasFocus) {
            super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
            if (value == null) {
                return this;
            }
            final var sheetName = (SheetName)value;
            final var label = sheetName.label();
            if (sheetName.isMissing()) {
                // mimics how missing columns are displayed in the webui twinlist
                setText("(MISSING) %s".formatted(label));
                setForeground(Color.RED);
            } else {
                setText(label);
                setForeground(Color.BLACK);
            }
            return this;
        }

        /**
         * Creates a handler that colors the foreground according to the "missing" status of the selected
         * {@link SheetName}.
         *
         * @param comboBox combobox to get selected item from and register handler with
         */
        private static void registerComboBoxSelectedMissingHandler(final JComboBox<SheetName> comboBox) {
            // this makes sure that a missing item, if selected, is visibly marked even when the dropdown is closed
            // (this does not alter the appearance when the combobox is disabled)
            comboBox.addActionListener(l -> {
                final var curr = (SheetName)comboBox.getSelectedItem();
                if (curr == null) {
                    return;
                }
                comboBox.setForeground(curr.isMissing() ? Color.RED : Color.BLACK);
                return;
            });
        }
    }

}
