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
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.Inside;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.text.TextInputWidget;

/**
 * Parameters for configuring how values are read from Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(ValuesParameters.ValuesLayout.class)
class ValuesParameters implements NodeParameters {

    @Inside(ExcelTableReaderParameters.ValuesSection.class)
    interface ValuesLayout {
    }

    static final class Use15DigitsPrecisionRef implements ParameterReference<Boolean> {
    }

    static final class ReplaceEmptyStringsWithMissingValuesRef implements ParameterReference<Boolean> {
    }

    static final class ReevaluateFormulasRef implements ParameterReference<Boolean> {
    }

    static final class FormulaErrorHandlingRef implements ParameterReference<FormulaErrorHandlingOption> {
    }

    static final class FormulaErrorPatternRef implements ParameterReference<String> {
    }

    private enum FormulaErrorHandlingOption {
            @Label(value = "String value",
                description = "Insert the given String value. Inserting a string value causes the entire column "
                    + "to become a string column in case an error occurs.")
            PATTERN,

            @Label(value = "Missing value",
                description = "Insert a Missing value. Inserting a missing value is type innocent, but also "
                    + "unobtrusive.")
            MISSING
    }

    private static final class IsFormulaErrorPatternEnabled implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(FormulaErrorHandlingRef.class).isOneOf(FormulaErrorHandlingOption.PATTERN);
        }
    }

    @Widget(title = "Use Excel 15 digits precision",
        description = "If checked, numbers are read in with 15 digits precision which is the same precision Excel is "
            + "using to display numbers. This will prevent potential floating point issues. For most numbers, "
            + "no difference can be observed if this option is unchecked.")
    @ValueReference(Use15DigitsPrecisionRef.class)
    boolean m_use15DigitsPrecision = true;

    @Widget(title = "Replace empty strings with missing values",
        description = "If checked, empty strings (i.e. strings with only whitespaces) are replaced with missing "
            + "values. This option is also applied to formulas that evaluate to strings.")
    @ValueReference(ReplaceEmptyStringsWithMissingValuesRef.class)
    boolean m_replaceEmptyStringsWithMissingValues = true;

    @Widget(title = "Reevaluate formulas in all sheets",
        description = "If checked, formulas are reevaluated and put into the created table instead of using the "
            + "cached values. This can cause errors when there are functions that are not implemented by the "
            + "underlying Apache POI library. Note: Files with xlsb format do not support this option. For xlsx and "
            + "xlsm files, reevaluation requires significantly more memory as the whole file needs to be kept in "
            + "memory (xls files are anyway loaded completely into memory).")
    @ValueReference(ReevaluateFormulasRef.class)
    boolean m_reevaluateFormulas;

    @Widget(title = "On error insert", description = "Choose how to handle formula evaluation errors.")
    @ValueSwitchWidget
    @ValueReference(FormulaErrorHandlingRef.class)
    FormulaErrorHandlingOption m_formulaErrorHandling = FormulaErrorHandlingOption.PATTERN;

    @Widget(title = "String", description = "The string value to insert into a cell with a formula evaluation error.")
    @TextInputWidget
    @ValueReference(FormulaErrorPatternRef.class)
    @Effect(predicate = IsFormulaErrorPatternEnabled.class, type = EffectType.SHOW)
    String m_formulaErrorPattern = ExcelTableReaderConfig.DEFAULT_FORMULA_ERROR_PATTERN;

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        excelConfig.setUse15DigitsPrecision(m_use15DigitsPrecision);
        excelConfig.setReplaceEmptyStringsWithMissings(m_replaceEmptyStringsWithMissingValues);
        excelConfig.setReevaluateFormulas(m_reevaluateFormulas);

        switch (m_formulaErrorHandling) {
            case PATTERN:
                excelConfig.setFormulaErrorHandling(FormulaErrorHandling.PATTERN);
                excelConfig.setErrorPattern(m_formulaErrorPattern);
                break;
            case MISSING:
                excelConfig.setFormulaErrorHandling(FormulaErrorHandling.MISSING);
                break;
        }
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        m_use15DigitsPrecision = excelConfig.isUse15DigitsPrecision();
        m_replaceEmptyStringsWithMissingValues = excelConfig.isReplaceEmptyStringsWithMissings();
        m_reevaluateFormulas = excelConfig.isReevaluateFormulas();
        m_formulaErrorPattern = excelConfig.getErrorPattern();

        switch (excelConfig.getFormulaErrorHandling()) {
            case PATTERN:
                m_formulaErrorHandling = FormulaErrorHandlingOption.PATTERN;
                break;
            case MISSING:
                m_formulaErrorHandling = FormulaErrorHandlingOption.MISSING;
                break;
            default:
                throw new IllegalStateException(
                    "Unknown FormulaErrorHandling: " + excelConfig.getFormulaErrorHandling());
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // no specific validation needed
    }

}
