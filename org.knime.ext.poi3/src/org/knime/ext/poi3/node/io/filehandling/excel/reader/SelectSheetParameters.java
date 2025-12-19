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
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.webui.node.dialog.defaultdialog.util.updates.StateComputationFailureException;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.StateProvider;
import org.knime.node.parameters.updates.ValueProvider;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.updates.legacy.AutoGuessValueProvider;
import org.knime.node.parameters.widget.choices.ChoicesProvider;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.StringChoicesProvider;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.number.NumberInputWidget;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation.MinValidation.IsNonNegativeValidation;

import java.util.List;
import java.util.Optional;
import java.util.function.Supplier;

/**
 * Parameters for sheet selection in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(SelectSheetParameters.SelectSheetLayout.class)
class SelectSheetParameters implements NodeParameters {

    @After(EncryptionParameters.EncryptionLayout.class)
    interface SelectSheetLayout {
    }

    static final class SheetSelectionRef implements ParameterReference<SheetSelectionMode> {
    }

    static final class SheetNameRef implements ParameterReference<String> {
    }

    static final class SheetIndexRef implements ParameterReference<Integer> {
    }

    private enum SheetSelectionMode {
            @Label(value = "First sheet with data",
                description = "The first sheet of the selected file(s) that contains data will be read in. "
                    + "Containing data means not being empty. If all sheets of a file are empty, an empty table is "
                    + "read in.")
            FIRST,

            @Label(value = "By name",
                description = "The sheet with the selected name will be read in. If reading multiple files, "
                    + "the sheet names of the first file are shown and the node will fail if any of the other files "
                    + "does not contain a sheet with the selected name.")
            NAME,

            @Label(value = "By position",
                description = "The sheet at the selected position will be read in. If reading multiple files, "
                    + "the node will fail if any of the files does not contain a sheet at the selected position. "
                    + "The position starts at 0, i.e. the first sheet is at position 0.")
            INDEX
    }

    private static final class IsSheetNameMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(SheetSelectionRef.class).isOneOf(SheetSelectionMode.NAME);
        }
    }

    private static final class IsSheetIndexMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(SheetSelectionRef.class).isOneOf(SheetSelectionMode.INDEX);
        }
    }

    @Widget(title = "Select sheet",
        description = "Choose which sheet to read from the Excel file. "
            + "The order of the sheets is the same as displayed in Excel (i.e. not necessarily a lexicographic order).")
    @ValueSwitchWidget
    @ValueReference(SheetSelectionRef.class)
    SheetSelectionMode m_sheetSelection = SheetSelectionMode.FIRST;

    @Widget(title = "Sheet name", description = "The name of the sheet to read.")
    @ValueReference(SheetNameRef.class)
    @Effect(predicate = IsSheetNameMode.class, type = EffectType.SHOW)
    @ChoicesProvider(SheetNamesChoicesProvider.class)
    @ValueProvider(SheetNamesDefaultProvider.class)
    String m_sheetName = "";

    @Widget(title = "Sheet position", description = "The position (starting at 0) of the sheet to read. "
        + "The maximum position that can be selected depends on the number of sheets available in the first read file.")
    @NumberInputWidget(minValidation = IsNonNegativeValidation.class,
        maxValidationProvider = MaxSheetIndexProvider.class)
    @ValueReference(SheetIndexRef.class)
    @Effect(predicate = IsSheetIndexMode.class, type = EffectType.SHOW)
    int m_sheetIndex;

    private static class SheetNamesChoicesProvider implements StringChoicesProvider {

        Supplier<ExcelFileContentInfoStateProvider.ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProviderInitializer initializer) {
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelFileContentInfoStateProvider.class);
        }

        @Override
        public List<String> choices(final NodeParametersInput context) {
            return m_excelFileInfo.get().getSheetNames();
        }
    }

    private static class SheetNamesDefaultProvider extends AutoGuessValueProvider<String> {
        SheetNamesDefaultProvider() {
            super(SheetNameRef.class);
        }

        Supplier<ExcelFileContentInfoStateProvider.ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProviderInitializer initializer) {
            super.init(initializer);
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelFileContentInfoStateProvider.class);
        }

        @Override
        protected boolean isEmpty(String value) {
            return value == null || value.isEmpty();
        }

        @Override
        protected String autoGuessValue(NodeParametersInput parametersInput) throws StateComputationFailureException {
            return Optional.of(m_excelFileInfo.get()) //
                .orElseThrow(StateComputationFailureException::new) //
                .getSheetNames().stream().findFirst() //
                .orElseThrow(StateComputationFailureException::new);
        }
    }

    private static class MaxSheetIndexProvider implements StateProvider<NumberInputWidgetValidation.MaxValidation> {

        Supplier<ExcelFileContentInfoStateProvider.ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProviderInitializer initializer) {
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelFileContentInfoStateProvider.class);
        }

        @Override
        public NumberInputWidgetValidation.MaxValidation computeState(NodeParametersInput parametersInput) {
            return new NumberInputWidgetValidation.MaxValidation() {
                @Override
                protected double getMax() {
                    return numberOfSheets() - 1;
                }

                private int numberOfSheets() {
                    return Math.max(m_excelFileInfo.get().getSheetNames().size(), 1);
                }

                @Override
                public String getErrorMessage() {
                    if (numberOfSheets() == 1) {
                        return "Only 1 sheet is available, therefore the only allowed sheet position is 0.";
                    }
                    return "Only %d sheets are available, therefore the maximum allowed sheet position is %d."
                        .formatted(numberOfSheets(), numberOfSheets() - 1);
                }
            };
        }
    }

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        switch (m_sheetSelection) {
            case FIRST:
                excelConfig.setSheetSelection(SheetSelection.FIRST);
                break;
            case NAME:
                excelConfig.setSheetSelection(SheetSelection.NAME);
                excelConfig.setSheetName(m_sheetName);
                break;
            case INDEX:
                excelConfig.setSheetSelection(SheetSelection.INDEX);
                excelConfig.setSheetIdx(m_sheetIndex);
                break;
        }
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        m_sheetName = excelConfig.getSheetName();
        m_sheetIndex = excelConfig.getSheetIdx();
        switch (excelConfig.getSheetSelection()) {
            case FIRST:
                m_sheetSelection = SheetSelectionMode.FIRST;
                break;
            case NAME:
                m_sheetSelection = SheetSelectionMode.NAME;
                break;
            case INDEX:
                m_sheetSelection = SheetSelectionMode.INDEX;
                break;
            default:
                throw new IllegalStateException("Unknown SheetSelection: " + excelConfig.getSheetSelection());
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // nothing to validate here
    }

}
