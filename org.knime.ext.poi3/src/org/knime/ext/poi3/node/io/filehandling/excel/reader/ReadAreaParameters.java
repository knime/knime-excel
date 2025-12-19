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

import java.util.ArrayList;
import java.util.Optional;
import java.util.function.Supplier;

import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
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
import org.knime.node.parameters.updates.StateProvider;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.message.TextMessage;
import org.knime.node.parameters.widget.text.TextInputWidget;
import org.knime.node.parameters.widget.text.TextInputWidgetValidation;

/**
 * Parameters for defining the area of the sheet to read in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(ReadAreaParameters.ReadAreaLayout.class)
class ReadAreaParameters implements NodeParameters {

    /**
     * Needs to be consistent with the regex defined in {@link ExcelUtils}, but allows empty string for first / last
     * column, or a numeric specification.
     */
    private static final String EXCEL_COLUMN_NAME_REGEX = "[A-Z]*|\\d*";

    /**
     * Number, or empty string for first / last row.
     */
    private static final String EXCEL_ROW_INDEX_REGEX = "\\d*";

    @Inside(ReaderLayout.DataArea.class)
    interface ReadAreaLayout {

        @HorizontalLayout
        @Effect(predicate = IsCustomAreaMode.class, type = EffectType.SHOW)
        interface ColumnSelection {
        }

        @HorizontalLayout
        @After(ColumnSelection.class)
        @Effect(predicate = IsCustomAreaMode.class, type = EffectType.SHOW)
        interface RowSelection {
        }

        @After(RowSelection.class)
        interface Errors {
        }

    }

    static final class SheetAreaRef implements ParameterReference<SheetAreaMode> {
    }

    static final class ReadFromColumnRef implements ParameterReference<String> {
    }

    static final class ReadToColumnRef implements ParameterReference<String> {
    }

    static final class ReadFromRowRef implements ParameterReference<String> {
    }

    static final class ReadToRowRef implements ParameterReference<String> {
    }

    private enum SheetAreaMode {
            @Label(value = "Whole sheet",
                description = "All the data contained in the sheet is read in. This includes areas where diagrams, "
                    + "borders, coloring, etc. are placed and could create empty rows or columns.")
            ENTIRE,

            @Label(value = "Range by column and row",
                description = "Only the data in the specified area is read in. Both start and end columns/rows are "
                    + "inclusive. By leaving a field empty, the start or end of the area is not restricted.")
            CUSTOM
    }

    private static final class IsCustomAreaMode implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(SheetAreaRef.class).isOneOf(SheetAreaMode.CUSTOM);
        }
    }

    @Widget(title = "Define read area", description = "Choose whether to read the entire sheet or a custom area.")
    @ValueSwitchWidget
    @ValueReference(SheetAreaRef.class)
    SheetAreaMode m_sheetArea = SheetAreaMode.ENTIRE;

    @Widget(title = "Read from column",
        description = "The first column to include specified as label (\"A\", \"B\", etc.) or number "
            + "(starting at 1). Leave empty to read starting at the first column.")
    @ValueReference(ReadFromColumnRef.class)
    @TextInputWidget(patternValidation = ExcelColumnNameValidation.class)
    @Layout(ReadAreaLayout.ColumnSelection.class)
    String m_readFromColumn = "A";

    @Widget(title = "Read to column",
        description = "The last column to include specified as label (\"A\", \"B\", etc.) or number "
            + "(starting at 1). Leave empty to read until the last column.")
    @ValueReference(ReadToColumnRef.class)
    @TextInputWidget(patternValidation = ExcelColumnNameValidation.class)
    @Layout(ReadAreaLayout.ColumnSelection.class)
    String m_readToColumn = "";

    @Widget(title = "Read from row",
        description = "The first row number to include (starting at 1). "
            + "Leave empty to read starting from the first row.")
    @ValueReference(ReadFromRowRef.class)
    @TextInputWidget(patternValidation = ExcelRowIndexValidation.class)
    @Layout(ReadAreaLayout.RowSelection.class)
    String m_readFromRow = "1";

    @Widget(title = "Read to row",
        description = "The last row number to include (starting at 1). Leave empty to read until the last row.")
    @ValueReference(ReadToRowRef.class)
    @TextInputWidget(patternValidation = ExcelRowIndexValidation.class)
    @Layout(ReadAreaLayout.RowSelection.class)
    String m_readToRow = "";

    @TextMessage(ValidationMessageProvider.class)
    @Layout(ReadAreaLayout.Errors.class)
    Void m_validationMessage;

    private static class ExcelColumnNameValidation extends TextInputWidgetValidation.PatternValidation {
        @Override
        protected String getPattern() {
            return EXCEL_COLUMN_NAME_REGEX;
        }

        @Override
        public String getErrorMessage() {
            return "Value must be an Excel column identifier (A, B, ... or 1, 2, ...) or empty";
        }
    }

    private static class ExcelRowIndexValidation extends TextInputWidgetValidation.PatternValidation {
        @Override
        protected String getPattern() {
            return EXCEL_ROW_INDEX_REGEX;
        }

        @Override
        public String getErrorMessage() {
            return "Value must be a positive integer or empty";
        }
    }

    private static class ValidationMessageProvider implements StateProvider<Optional<TextMessage.Message>> {

        private Supplier<SheetAreaMode> m_sheetArea;

        private Supplier<String> m_readFromColumn;

        private Supplier<String> m_readToColumn;

        private Supplier<String> m_readFromRow;

        private Supplier<String> m_readToRow;

        @Override
        public void init(StateProviderInitializer initializer) {
            initializer.computeBeforeOpenDialog();
            m_sheetArea = initializer.computeFromValueSupplier(SheetAreaRef.class);
            m_readFromColumn = initializer.computeFromValueSupplier(ReadFromColumnRef.class);
            m_readToColumn = initializer.computeFromValueSupplier(ReadToColumnRef.class);
            m_readFromRow = initializer.computeFromValueSupplier(ReadFromRowRef.class);
            m_readToRow = initializer.computeFromValueSupplier(ReadToRowRef.class);
        }

        @Override
        public Optional<TextMessage.Message> computeState(NodeParametersInput parametersInput) {
            var errors = new ArrayList<String>();

            if (m_sheetArea.get() == SheetAreaMode.CUSTOM) {
                try {
                    ExcelUtils.validateColIndexes(m_readFromColumn.get(), m_readToColumn.get());
                } catch (InvalidSettingsException e) { // NOSONAR swallowing exception is intentional
                    errors.add(e.getMessage());
                } catch (IllegalArgumentException e) { // NOSONAR swallowing exception is intentional
                    // ignore, as this is already handled by the pattern validation
                }
                try {
                    ExcelUtils.validateRowIndexes(m_readFromRow.get(), m_readToRow.get());
                } catch (InvalidSettingsException e) { // NOSONAR swallowing exception is intentional
                    errors.add(e.getMessage());
                } catch (IllegalArgumentException e) { // NOSONAR swallowing exception is intentional
                    // ignore, as this is already handled by the pattern validation
                }
            }

            if (!errors.isEmpty()) {
                String description = String.join("\n", errors);
                return Optional.of(
                    new TextMessage.Message("Invalid read area definition", description, TextMessage.MessageType.INFO));
            }

            return Optional.empty();
        }
    }

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        switch (m_sheetArea) {
            case ENTIRE:
                excelConfig.setAreaOfSheetToRead(AreaOfSheetToRead.ENTIRE);
                break;
            case CUSTOM:
                excelConfig.setAreaOfSheetToRead(AreaOfSheetToRead.PARTIAL);
                excelConfig.setReadFromCol(m_readFromColumn);
                excelConfig.setReadToCol(m_readToColumn);
                excelConfig.setReadFromRow(m_readFromRow);
                excelConfig.setReadToRow(m_readToRow);
                break;
        }
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();

        m_readFromColumn = excelConfig.getReadFromCol();
        m_readToColumn = excelConfig.getReadToCol();
        m_readFromRow = excelConfig.getReadFromRow();
        m_readToRow = excelConfig.getReadToRow();
        switch (excelConfig.getAreaOfSheetToRead()) {
            case ENTIRE:
                m_sheetArea = SheetAreaMode.ENTIRE;
                break;
            case PARTIAL:
                m_sheetArea = SheetAreaMode.CUSTOM;
                break;
            default:
                throw new IllegalStateException("Unknown AreaOfSheetToRead: " + excelConfig.getAreaOfSheetToRead());
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        if (m_sheetArea == SheetAreaMode.CUSTOM) {
            ExcelUtils.validateColIndexes(m_readFromColumn, m_readToColumn);
            ExcelUtils.validateRowIndexes(m_readFromRow, m_readToRow);
        }
    }

}
