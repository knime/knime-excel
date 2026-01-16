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

import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelMultiTableReadConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.RowIDGeneration;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
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
import org.knime.node.parameters.widget.text.TextInputWidgetValidation.PatternValidation;

/**
 * Parameters for defining RowID in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Advanced
@Layout(RowIdParameters.RowIdLayout.class)
class RowIdParameters implements NodeParameters {

    @Inside(ReaderLayout.DataArea.class)
    @After(ColumnNameParameters.ColumnNameLayout.class)
    interface RowIdLayout {
    }

    static final class RowIdModeRef implements ParameterReference<RowIdMode> {
    }

    static final class RowIdColumnRef implements ParameterReference<String> {
    }

    private enum RowIdMode {
            @Label(value = "Enumerate rows (Row1, Row2, ...)", description = "Generate row IDs by enumerating rows.")
            ENUMERATE,

            @Label(value = "Use custom column", description = "Use a specific column as row IDs.")
            CUSTOM_COLUMN
    }

    private static final class IsCustomColumnMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(RowIdModeRef.class).isOneOf(RowIdMode.CUSTOM_COLUMN);
        }
    }

    private static final class ExcelColumnPatternValidation extends PatternValidation {
        @Override
        protected String getPattern() {
            // Matches Excel column letters or numeric column indices (starting at 1)
            return "[A-Z]+|[1-9][0-9]*";
        }

        @Override
        public String getErrorMessage() {
            return "The value must be a valid Excel column identifier (e.g., A, B, AA, AB or 1, 2, 3).";
        }
    }

    @Widget(title = "Define RowID", description = "Choose how to define row IDs for the table.")
    @ValueSwitchWidget
    @ValueReference(RowIdModeRef.class)
    RowIdMode m_rowIdMode = RowIdMode.ENUMERATE;

    @Widget(title = "RowID as in column", description = "The column (e.g., A, B, C or 1, 2, 3) to use as row IDs.")
    @TextInputWidget(patternValidation = ExcelColumnPatternValidation.class)
    @ValueReference(RowIdColumnRef.class)
    @Effect(predicate = IsCustomColumnMode.class, type = EffectType.SHOW)
    String m_rowIdColumn = "A";

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var tableReadConfig = config.getTableReadConfig();
        final var excelConfig = config.getReaderSpecificConfig();

        switch (m_rowIdMode) {
            // TODO is it correct to store those settings both into excelConfig and tableReadConfig?
            // TODO it seems not to work if storing them only into the tableReadConfig.
            case ENUMERATE:
                excelConfig.setRowIdGeneration(RowIDGeneration.GENERATE);
                tableReadConfig.setUseRowIDIdx(false);
                break;
            case CUSTOM_COLUMN:
                excelConfig.setRowIDCol(m_rowIdColumn);
                excelConfig.setRowIdGeneration(RowIDGeneration.COLUMN);
                tableReadConfig.setUseRowIDIdx(true);
                tableReadConfig.setRowIDIdx(ExcelUtils.getRowIdColIdx(m_rowIdColumn));
                break;
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // no specific validation needed
    }

}
