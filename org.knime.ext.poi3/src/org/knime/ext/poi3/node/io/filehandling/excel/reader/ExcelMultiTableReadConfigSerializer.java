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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.SettingsModel;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelReadAdapterFactory;
import org.knime.filehandling.core.node.table.reader.config.ConfigSerializer;
import org.knime.filehandling.core.node.table.reader.config.DefaultMultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableSpecConfig;
import org.knime.filehandling.core.util.SettingsUtils;

/**
 * {@link ConfigSerializer} for the Excel reader node.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
enum ExcelMultiTableReadConfigSerializer implements
    ConfigSerializer<DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>>> {

        /**
         * Singleton instance.
         */
        INSTANCE;

    private static final KNIMECellType MOST_GENERIC_EXTERNAL_TYPE = KNIMECellType.STRING;

    private static final String CFG_TABLE_SPEC_CONFIG = "table_spec_config" + SettingsModel.CFGKEY_INTERNAL;

    private static final String CFG_FAIL_ON_DIFFERING_SPECS = "fail_on_differing_specs";

    private static final String CFG_SETTINGS_TAB = "settings";

    private static final String CFG_SHEET_SELECTION = "sheet_selection";

    private static final String CFG_SHEET_NAME = "sheet_name";

    private static final String CFG_SHEET_INDEX = "sheet_index";

    private static final String CFG_HAS_COLUMN_HEADER = "table_contains_column_names";

    private static final String CFG_COLUMN_HEADER_ROW_NUMBER = "column_names_row_number";

    private static final String CFG_ADVANCED_SETTINGS_TAB = "advanced_settings";

    private static final String CFG_USE_15_DIGITS_PRECISION = "use_15_digits_precision";

    private static final String CFG_SKIP_HIDDEN_COLS = "skip_hidden_columns";

    private static final String CFG_SKIP_HIDDEN_ROWS = "skip_hidden_rows";

    private static final String CFG_SKIP_EMPTY_ROWS = "skip_empty_rows";

    private static final String CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS = "replace_empty_strings_with_missings";

    private static final String CFG_REEVALUATE_FORMULAS = "reevaluate_formulas";

    private static final String CFG_FORMULA_ERROR_HANDLING = "formula_error_handling";

    private static final String CFG_FORMULA_ERROR_PATTERN = "formula_error_pattern";

    private static final String CFG_AREA_OF_SHEET_TO_READ = "sheet_area";

    private static final String CFG_READ_FROM_COL = "read_from_column";

    private static final String CFG_READ_TO_COL = "read_to_column";

    private static final String CFG_READ_FROM_ROW = "read_from_row";

    private static final String CFG_READ_TO_ROW = "read_to_row";

    private static final String CFG_ROW_ID_GENERATION = "row_id";

    private static final String CFG_ROW_ID_COL = "row_id_column";

    private static final String CFG_MAX_DATA_ROWS_SCANNED = "max_data_rows_scanned";

    private static final String CFG_LIMIT_DATA_ROWS_SCANNED = "limit_data_rows_scanned";

    @Override
    public void loadInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings, final PortObjectSpec[] specs) throws NotConfigurableException {
        loadSettingsTabInDialog(config, SettingsUtils.getOrEmpty(settings, CFG_SETTINGS_TAB));
        loadAdvancedSettingsTabInDialog(config, SettingsUtils.getOrEmpty(settings, CFG_ADVANCED_SETTINGS_TAB));
        if (settings.containsKey(CFG_TABLE_SPEC_CONFIG)) {
            try {
                config.setTableSpecConfig(DefaultTableSpecConfig.load(settings.getNodeSettings(CFG_TABLE_SPEC_CONFIG),
                    ExcelReadAdapterFactory.INSTANCE.getProducerRegistry(), MOST_GENERIC_EXTERNAL_TYPE, null));
            } catch (InvalidSettingsException ex) { // NOSONAR
                /* Can only happen in TableSpecConfig#load, since we checked #NodeSettingsRO#getNodeSettings(String)
                 * before. The framework takes care that #validate is called before load so we can assume that this
                 * exception does not occur.
                 */
            }
        } else {
            config.setTableSpecConfig(null);
        }
    }

    @Override
    public void loadInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        loadSettingsTabInModel(config, SettingsUtils.getOrEmpty(settings, CFG_SETTINGS_TAB));
        loadAdvancedSettingsTabInModel(config, settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        if (settings.containsKey(CFG_TABLE_SPEC_CONFIG)) {
            config.setTableSpecConfig(DefaultTableSpecConfig.load(settings.getNodeSettings(CFG_TABLE_SPEC_CONFIG),
                ExcelReadAdapterFactory.INSTANCE.getProducerRegistry(), MOST_GENERIC_EXTERNAL_TYPE, null));
        } else {
            config.setTableSpecConfig(null);
        }
    }

    @Override
    public void saveInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        if (config.hasTableSpecConfig()) {
            config.getTableSpecConfig().save(settings.addNodeSettings(CFG_TABLE_SPEC_CONFIG));
        }
        saveSettingsTab(config, SettingsUtils.getOrAdd(settings, CFG_SETTINGS_TAB));
        saveAdvancedSettingsTab(config, settings.addNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
    }

    @Override
    public void validate(final NodeSettingsRO settings) throws InvalidSettingsException {
        validateSettingsTab(settings.getNodeSettings(CFG_SETTINGS_TAB));
        validateAdvancedSettingsTab(settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
    }

    @Override
    public void saveInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) throws InvalidSettingsException {
        saveInModel(config, settings);
    }

    private static void loadSettingsTabInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setSheetSelection(
            SheetSelection.valueOf(settings.getString(CFG_SHEET_SELECTION, SheetSelection.FIRST.name())));
        excelConfig.setSheetName(settings.getString(CFG_SHEET_NAME, ""));
        excelConfig.setSheetIdx(settings.getInt(CFG_SHEET_INDEX, 0));
        excelConfig.setRowIdGeneration(
            RowIDGeneration.valueOf(settings.getString(CFG_ROW_ID_GENERATION, RowIDGeneration.GENERATE.name())));
        excelConfig.setRowIDCol(settings.getString(CFG_ROW_ID_COL, "A"));
        excelConfig.setAreaOfSheetToRead(
            AreaOfSheetToRead.valueOf(settings.getString(CFG_AREA_OF_SHEET_TO_READ, AreaOfSheetToRead.ENTIRE.name())));
        excelConfig.setReadFromCol(settings.getString(CFG_READ_FROM_COL, "A"));
        excelConfig.setReadToCol(settings.getString(CFG_READ_TO_COL, ""));
        excelConfig.setReadFromRow(settings.getString(CFG_READ_FROM_ROW, "1"));
        excelConfig.setReadToRow(settings.getString(CFG_READ_TO_ROW, ""));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER, true));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_NUMBER, 1) - 1);
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(false);
    }

    private static void loadSettingsTabInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setSheetSelection(SheetSelection.loadValueInModel(settings.getString(CFG_SHEET_SELECTION)));
        excelConfig.setSheetName(settings.getString(CFG_SHEET_NAME));
        excelConfig.setSheetIdx(settings.getInt(CFG_SHEET_INDEX));
        excelConfig
            .setAreaOfSheetToRead(AreaOfSheetToRead.loadValueInModel(settings.getString(CFG_AREA_OF_SHEET_TO_READ)));
        excelConfig.setRowIdGeneration(RowIDGeneration.valueOf(settings.getString(CFG_ROW_ID_GENERATION)));
        excelConfig.setRowIDCol(settings.getString(CFG_ROW_ID_COL));
        excelConfig.setReadFromCol(settings.getString(CFG_READ_FROM_COL));
        excelConfig.setReadToCol(settings.getString(CFG_READ_TO_COL));
        excelConfig.setReadFromRow(settings.getString(CFG_READ_FROM_ROW));
        excelConfig.setReadToRow(settings.getString(CFG_READ_TO_ROW));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_NUMBER) - 1);
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(true);
        tableReadConfig.setUseRowIDIdx(excelConfig.getRowIdGeneration() == RowIDGeneration.COLUMN);
        tableReadConfig.setRowIDIdx(0);
        tableReadConfig.setDecorateRead(false);
    }

    private static void saveSettingsTab(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addString(CFG_SHEET_SELECTION, excelConfig.getSheetSelection().name());
        settings.addString(CFG_SHEET_NAME, excelConfig.getSheetName());
        settings.addInt(CFG_SHEET_INDEX, excelConfig.getSheetIdx());
        settings.addBoolean(CFG_HAS_COLUMN_HEADER, tableReadConfig.useColumnHeaderIdx());
        settings.addLong(CFG_COLUMN_HEADER_ROW_NUMBER, tableReadConfig.getColumnHeaderIdx() + 1);
        settings.addString(CFG_ROW_ID_GENERATION, excelConfig.getRowIdGeneration().name());
        settings.addString(CFG_ROW_ID_COL, excelConfig.getRowIDCol());
        settings.addString(CFG_AREA_OF_SHEET_TO_READ, excelConfig.getAreaOfSheetToRead().name());
        settings.addString(CFG_READ_FROM_COL, excelConfig.getReadFromCol());
        settings.addString(CFG_READ_TO_COL, excelConfig.getReadToCol());
        settings.addString(CFG_READ_FROM_ROW, excelConfig.getReadFromRow());
        settings.addString(CFG_READ_TO_ROW, excelConfig.getReadToRow());
    }

    public static void validateSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
        SheetSelection.loadValueInModel(settings.getString(CFG_SHEET_SELECTION));
        settings.getString(CFG_SHEET_NAME);
        settings.getInt(CFG_SHEET_INDEX);
        settings.getBoolean(CFG_HAS_COLUMN_HEADER);
        settings.getLong(CFG_COLUMN_HEADER_ROW_NUMBER);
        RowIDGeneration.valueOf(settings.getString(CFG_ROW_ID_GENERATION));
        settings.getString(CFG_ROW_ID_COL);
        AreaOfSheetToRead.loadValueInModel(settings.getString(CFG_AREA_OF_SHEET_TO_READ));
        settings.getString(CFG_READ_FROM_COL);
        settings.getString(CFG_READ_TO_COL);
        settings.getString(CFG_READ_FROM_ROW);
        settings.getString(CFG_READ_TO_ROW);
    }

    private static void loadAdvancedSettingsTabInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) {
        config.setFailOnDifferingSpecs(settings.getBoolean(CFG_FAIL_ON_DIFFERING_SPECS, true));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setSkipEmptyRows(settings.getBoolean(CFG_SKIP_EMPTY_ROWS, true));
        tableReadConfig.setLimitRowsForSpec(settings.getBoolean(CFG_LIMIT_DATA_ROWS_SCANNED, true));
        tableReadConfig.setMaxRowsForSpec(settings.getLong(CFG_MAX_DATA_ROWS_SCANNED, 1000));
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CFG_USE_15_DIGITS_PRECISION, true));
        excelConfig.setSkipHiddenCols(settings.getBoolean(CFG_SKIP_HIDDEN_COLS, true));
        excelConfig.setSkipHiddenRows(settings.getBoolean(CFG_SKIP_HIDDEN_ROWS, true));
        excelConfig
            .setReplaceEmptyStringsWithMissings(settings.getBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS, true));
        excelConfig.setReevaluateFormulas(settings.getBoolean(CFG_REEVALUATE_FORMULAS, false));
        excelConfig.setFormulaErrorHandling(FormulaErrorHandling
            .valueOf(settings.getString(CFG_FORMULA_ERROR_HANDLING, FormulaErrorHandling.PATTERN.name())));
        excelConfig.setErrorPattern(
            settings.getString(CFG_FORMULA_ERROR_PATTERN, ExcelTableReaderConfig.DEFAULT_FORMULA_ERROR_PATTERN));
    }

    private static void loadAdvancedSettingsTabInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        config.setFailOnDifferingSpecs(settings.getBoolean(CFG_FAIL_ON_DIFFERING_SPECS));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setSkipEmptyRows(settings.getBoolean(CFG_SKIP_EMPTY_ROWS));
        tableReadConfig.setLimitRowsForSpec(settings.getBoolean(CFG_LIMIT_DATA_ROWS_SCANNED));
        tableReadConfig.setMaxRowsForSpec(settings.getLong(CFG_MAX_DATA_ROWS_SCANNED));
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CFG_USE_15_DIGITS_PRECISION));
        excelConfig.setSkipHiddenCols(settings.getBoolean(CFG_SKIP_HIDDEN_COLS));
        excelConfig.setSkipHiddenRows(settings.getBoolean(CFG_SKIP_HIDDEN_ROWS));
        excelConfig.setReplaceEmptyStringsWithMissings(settings.getBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS));
        excelConfig.setReevaluateFormulas(settings.getBoolean(CFG_REEVALUATE_FORMULAS));
        excelConfig.setFormulaErrorHandling(
            FormulaErrorHandling.loadValueInModel(settings.getString(CFG_FORMULA_ERROR_HANDLING)));
        excelConfig.setErrorPattern(settings.getString(CFG_FORMULA_ERROR_PATTERN));
    }

    private static void saveAdvancedSettingsTab(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        settings.addBoolean(CFG_FAIL_ON_DIFFERING_SPECS, config.failOnDifferingSpecs());
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        settings.addBoolean(CFG_SKIP_EMPTY_ROWS, tableReadConfig.skipEmptyRows());
        settings.addBoolean(CFG_LIMIT_DATA_ROWS_SCANNED, tableReadConfig.limitRowsForSpec());
        settings.addLong(CFG_MAX_DATA_ROWS_SCANNED, tableReadConfig.getMaxRowsForSpec());
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addBoolean(CFG_USE_15_DIGITS_PRECISION, excelConfig.isUse15DigitsPrecision());
        settings.addBoolean(CFG_SKIP_HIDDEN_COLS, excelConfig.isSkipHiddenCols());
        settings.addBoolean(CFG_SKIP_HIDDEN_ROWS, excelConfig.isSkipHiddenRows());
        settings.addBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS, excelConfig.isReplaceEmptyStringsWithMissings());
        settings.addBoolean(CFG_REEVALUATE_FORMULAS, excelConfig.isReevaluateFormulas());
        settings.addString(CFG_FORMULA_ERROR_HANDLING, excelConfig.getFormulaErrorHandling().name());
        settings.addString(CFG_FORMULA_ERROR_PATTERN, excelConfig.getErrorPattern());
    }

    public static void validateAdvancedSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
        settings.getBoolean(CFG_FAIL_ON_DIFFERING_SPECS);
        settings.getBoolean(CFG_USE_15_DIGITS_PRECISION);
        settings.getBoolean(CFG_SKIP_HIDDEN_COLS);
        settings.getBoolean(CFG_SKIP_HIDDEN_ROWS);
        settings.getBoolean(CFG_SKIP_EMPTY_ROWS);
        settings.getBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS);
        settings.getBoolean(CFG_REEVALUATE_FORMULAS);
        settings.getString(CFG_FORMULA_ERROR_HANDLING);
        settings.getString(CFG_FORMULA_ERROR_PATTERN);
        settings.getBoolean(CFG_LIMIT_DATA_ROWS_SCANNED);
        settings.getLong(CFG_MAX_DATA_ROWS_SCANNED);
    }

}
