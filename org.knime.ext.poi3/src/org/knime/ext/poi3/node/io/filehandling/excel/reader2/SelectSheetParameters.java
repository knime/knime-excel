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
package org.knime.ext.poi3.node.io.filehandling.excel.reader2;

import org.knime.core.node.InvalidSettingsException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelMultiTableReadConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.SheetSelection;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.number.NumberInputWidget;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation.MinValidation.IsNonNegativeValidation;
import org.knime.node.parameters.widget.text.TextInputWidget;

/**
 * Parameters for sheet selection in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(SelectSheetParameters.SelectSheetLayout.class)
class SelectSheetParameters implements NodeParameters {

    @After(ExcelTableReaderParameters.FileAndSheetSection.Source.class)
    interface SelectSheetLayout {
    }

    static final class SheetSelectionRef implements ParameterReference<SheetSelectionMode> {
    }

    static final class SheetNameRef implements ParameterReference<String> {
    }

    static final class SheetIndexRef implements ParameterReference<Integer> {
    }

    private enum SheetSelectionMode {
            @Label(value = "First sheet with data", description = "Select the first sheet that contains data.")
            FIRST,

            @Label(value = "By name", description = "Select a sheet by its name.")
            NAME,

            @Label(value = "By position", description = "Select a sheet by its position (0-based index).")
            INDEX,
    }

    static final class IsSheetNameMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(SheetSelectionRef.class).isOneOf(SheetSelectionMode.NAME);
        }
    }

    static final class IsSheetIndexMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(SheetSelectionRef.class).isOneOf(SheetSelectionMode.INDEX);
        }
    }

    @Widget(title = "Select sheet", description = "Choose how to select the sheet to read from the Excel file.")
    @ValueSwitchWidget
    @ValueReference(SheetSelectionRef.class)
    SheetSelectionMode m_sheetSelection = SheetSelectionMode.FIRST;

    // TODO provide choices based on excel file selected in file selection. Use a dropdown.
    @Widget(title = "Sheet name", description = "The name of the sheet to read.")
    @ValueReference(SheetNameRef.class)
    @Effect(predicate = IsSheetNameMode.class, type = EffectType.SHOW)
    @TextInputWidget
    String m_sheetName = "";

    @Widget(title = "Sheet position", description = "The position (0-based index) of the sheet to read.")
    @NumberInputWidget(minValidation = IsNonNegativeValidation.class)
    @ValueReference(SheetIndexRef.class)
    @Effect(predicate = IsSheetIndexMode.class, type = EffectType.SHOW)
    int m_sheetIndex = 0;

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        // Convert SheetSelectionMode to SheetSelection enum
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

    @Override
    public void validate() throws InvalidSettingsException {
        if (m_sheetSelection == SheetSelectionMode.NAME && (m_sheetName == null || m_sheetName.trim().isEmpty())) {
            throw new InvalidSettingsException("Sheet name must not be empty when selecting by name.");
        }
        if (m_sheetSelection == SheetSelectionMode.INDEX && m_sheetIndex < 0) {
            throw new InvalidSettingsException("Sheet position must be non-negative.");
        }
    }

}
