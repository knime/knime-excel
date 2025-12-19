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

import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderValidation.MaxNumberOfExcelRows;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.columnnames.ColumnNameMode;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.HorizontalLayout;
import org.knime.node.parameters.layout.Inside;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.RadioButtonsWidget;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.number.NumberInputWidget;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation.MinValidation.IsPositiveIntegerValidation;
import org.knime.node.parameters.widget.text.TextInputWidget;

/**
 * Parameters for defining column names in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Advanced
@Layout(ColumnNameParameters.ColumnNameLayout.class)
class ColumnNameParameters implements NodeParameters {

    @Inside(ReaderLayout.DataArea.class)
    @After(ReadAreaParameters.ReadAreaLayout.class)
    interface ColumnNameLayout {

        @Effect(predicate = IsCustomRowMode.class, type = EffectType.SHOW)
        @HorizontalLayout
        interface EmptyColumnFallback {
        }

    }

    static final class ColumnNameModeRef implements ParameterReference<ColumnNameModeOption> {
    }

    static final class ColumnNamesRowNumberRef implements ParameterReference<Long> {
    }

    static final class EmptyColHeaderPrefixRef implements ParameterReference<String> {
    }

    static final class EmptyColHeaderSuffixRef implements ParameterReference<EmptyColHeaderSuffixOption> {
    }

    private enum ColumnNameModeOption {
            @Label(value = "Use custom row", description = "Use a specific row as column names.")
            CUSTOM_ROW,

            @Label(value = "Excel headers (A, B, C)",
                description = "Use Excel column headers (A, B, C, ...) as column names.")
            EXCEL_HEADERS,

            @Label(value = "Enumerate columns (Col1, Col2...)",
                description = "Use enumerated column names (Col1, Col2, ...) as column names.")
            ENUMERATE_COLUMNS,
    }

    enum EmptyColHeaderSuffixOption {
            @Label(value = "A, B, ...",
                description = "Use Excel column headers (A, B, C, ...) as column name suffix for empty columns.")
            EXCEL_HEADERS,

            @Label(value = "1, 2, ...",
                description = "Use enumerated indices (1, 2, ...) as column name suffix for empty columns.")
            ENUMERATE_COLUMNS;

        /**
         * Converts this suffix option to the corresponding legacy ColumnNameMode.
         *
         * @return the ColumnNameMode for fallback in case of empty column headers
         */
        private ColumnNameMode toColumnNameMode() {
            return switch (this) {
                case EXCEL_HEADERS -> ColumnNameMode.EXCEL_COL_NAME;
                case ENUMERATE_COLUMNS -> ColumnNameMode.COL_INDEX;
            };
        }

        /**
         * Converts a legacy ColumnNameMode to the corresponding suffix option.
         *
         * @param mode the ColumnNameMode
         * @return the corresponding EmptyColHeaderSuffixOption
         */
        private static EmptyColHeaderSuffixOption fromColumnNameMode(final ColumnNameMode mode) {
            return switch (mode) {
                case EXCEL_COL_NAME -> EXCEL_HEADERS;
                case COL_INDEX -> ENUMERATE_COLUMNS;
            };
        }
    }

    private static final class IsCustomRowMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(ColumnNameModeRef.class).isOneOf(ColumnNameModeOption.CUSTOM_ROW);
        }
    }

    @Widget(title = "Define Column Names", description = "Choose how to define column names for the table.")
    @RadioButtonsWidget
    @ValueReference(ColumnNameModeRef.class)
    ColumnNameModeOption m_columnNameMode = ColumnNameModeOption.CUSTOM_ROW;

    @Widget(title = "Header as in row", description = "The row number (1-based) containing column names.")
    @NumberInputWidget(minValidation = IsPositiveIntegerValidation.class, maxValidation = MaxNumberOfExcelRows.class)
    @ValueReference(ColumnNamesRowNumberRef.class)
    @Effect(predicate = IsCustomRowMode.class, type = EffectType.SHOW)
    long m_columnNamesRowNumber = 1;

    @Widget(title = "Prefix for empty headers", description = "Prefix to use for empty column headers.")
    @TextInputWidget
    @ValueReference(EmptyColHeaderPrefixRef.class)
    @Layout(ColumnNameLayout.EmptyColumnFallback.class)
    String m_emptyColHeaderPrefix = "empty_";

    @Widget(title = "Suffix for empty headers",
        description = "Choose the suffix format for empty column headers.")
    @ValueSwitchWidget
    @ValueReference(EmptyColHeaderSuffixRef.class)
    @Layout(ColumnNameLayout.EmptyColumnFallback.class)
    EmptyColHeaderSuffixOption m_emptyColHeaderSuffix = EmptyColHeaderSuffixOption.EXCEL_HEADERS;

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var tableReadConfig = config.getTableReadConfig();

        switch (m_columnNameMode) {
            case EXCEL_HEADERS:
                excelConfig.setColumnNameMode(ColumnNameMode.EXCEL_COL_NAME);
                tableReadConfig.setUseColumnHeaderIdx(false);
                break;
            case ENUMERATE_COLUMNS:
                excelConfig.setColumnNameMode(ColumnNameMode.COL_INDEX);
                tableReadConfig.setUseColumnHeaderIdx(false);
                break;
            case CUSTOM_ROW:
                excelConfig.setColumnNameMode(m_emptyColHeaderSuffix.toColumnNameMode());
                tableReadConfig.setUseColumnHeaderIdx(true);
                tableReadConfig.setColumnHeaderIdx(m_columnNamesRowNumber - 1); // Convert to 0-based
                excelConfig.setEmptyColHeaderPrefix(m_emptyColHeaderPrefix);
                break;
        }
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var tableReadConfig = config.getTableReadConfig();

        m_columnNamesRowNumber = tableReadConfig.getColumnHeaderIdx() + 1; // Convert to 1-based
        m_emptyColHeaderPrefix = excelConfig.getEmptyColHeaderPrefix();
        if (tableReadConfig.useColumnHeaderIdx()) {
            m_columnNameMode = ColumnNameModeOption.CUSTOM_ROW;
            m_emptyColHeaderSuffix = EmptyColHeaderSuffixOption.fromColumnNameMode(excelConfig.getColumnNameMode());
        } else {
            switch (excelConfig.getColumnNameMode()) {
                case EXCEL_COL_NAME:
                    m_columnNameMode = ColumnNameModeOption.EXCEL_HEADERS;
                    break;
                case COL_INDEX:
                    m_columnNameMode = ColumnNameModeOption.ENUMERATE_COLUMNS;
                    break;
                default:
                    throw new IllegalStateException("Unknown ColumnNameMode: " + excelConfig.getColumnNameMode());
            }
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // no validations
    }

}
