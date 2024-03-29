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
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModel;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelReadAdapterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.columnnames.ColumnNameMode;
import org.knime.filehandling.core.node.table.ConfigSerializer;
import org.knime.filehandling.core.node.table.reader.config.AbstractTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.tablespec.ConfigID;
import org.knime.filehandling.core.node.table.reader.config.tablespec.ConfigIDFactory;
import org.knime.filehandling.core.node.table.reader.config.tablespec.DefaultProductionPathSerializer;
import org.knime.filehandling.core.node.table.reader.config.tablespec.NodeSettingsConfigID;
import org.knime.filehandling.core.node.table.reader.config.tablespec.NodeSettingsSerializer;
import org.knime.filehandling.core.node.table.reader.config.tablespec.TableSpecConfig;
import org.knime.filehandling.core.node.table.reader.config.tablespec.TableSpecConfigSerializer;
import org.knime.filehandling.core.util.SettingsUtils;

/**
 * {@link ConfigSerializer} for the Excel reader node.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
enum ExcelMultiTableReadConfigSerializer
    implements ConfigSerializer<ExcelMultiTableReadConfig>, ConfigIDFactory<ExcelMultiTableReadConfig> {

        INSTANCE;

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelMultiTableReadConfigSerializer.class);

    private static final String CFG_PATH_COLUMN_NAME = "path_column_name" + SettingsModel.CFGKEY_INTERNAL;

    private static final String CFG_APPEND_PATH_COLUMN = "append_path_column" + SettingsModel.CFGKEY_INTERNAL;

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

    private static final String CFG_SKIP_EMPTY_COLS = "skip_empty_cols";

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

    private static final String CFG_EMPTY_COL_HEADER_PREFIX = "empty_col_header_prefix";

    private static final String CFG_COLUMN_NAME_MODE = "column_name_mode";

    private static final String CFG_MAX_DATA_ROWS_SCANNED = "max_data_rows_scanned";

    private static final String CFG_LIMIT_DATA_ROWS_SCANNED = "limit_data_rows_scanned";

    private static final String CFG_SAVE_TABLE_SPEC_CONFIG = "save_table_spec_config" + SettingsModel.CFGKEY_INTERNAL;

    private static final String CFG_CHECK_TABLE_SPEC = "check_table_spec";

    private static final String CFG_ENCRYPTION_SETTINGS_TAB = "encryption";

    private final TableSpecConfigSerializer<KNIMECellType> m_tableSpecConfigSerializer;

    private enum ExcelTypeSerializer implements NodeSettingsSerializer<KNIMECellType> {
            SERIALIZER;

        private static final String CFG_TYPE = "type";

        @Override
        public void save(final KNIMECellType object, final NodeSettingsWO settings) {
            settings.addString(CFG_TYPE, object.name());
        }

        @Override
        public KNIMECellType load(final NodeSettingsRO settings) throws InvalidSettingsException {
            return KNIMECellType.valueOf(settings.getString(CFG_TYPE));
        }

    }

    ExcelMultiTableReadConfigSerializer() {
        m_tableSpecConfigSerializer = TableSpecConfigSerializer.createStartingV43(
            new DefaultProductionPathSerializer(ExcelReadAdapterFactory.INSTANCE.getProducerRegistry()), this,
            ExcelTypeSerializer.SERIALIZER);
    }

    @Override
    public void loadInDialog(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings,
        final PortObjectSpec[] specs) {
        loadSettingsTabInDialog(config, SettingsUtils.getOrEmpty(settings, CFG_SETTINGS_TAB));
        loadAdvancedSettingsTabInDialog(config, SettingsUtils.getOrEmpty(settings, CFG_ADVANCED_SETTINGS_TAB));
        try {
            config.setTableSpecConfig(loadTableSpecConfig(settings));
        } catch (InvalidSettingsException ex) { // NOSONAR
            /* Can only happen in TableSpecConfig#load, since we checked #NodeSettingsRO#getNodeSettings(String)
             * before. The framework takes care that #validate is called before load so we can assume that this
             * exception does not occur.
             */
        }
    }

    private TableSpecConfig<KNIMECellType> loadTableSpecConfig(final NodeSettingsRO settings)
        throws InvalidSettingsException {
        if (settings.containsKey(CFG_TABLE_SPEC_CONFIG)) {
            return m_tableSpecConfigSerializer.load(settings.getNodeSettings(CFG_TABLE_SPEC_CONFIG));
        } else {
            return null;
        }
    }

    @Override
    public void loadInModel(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings)
        throws InvalidSettingsException {
        loadSettingsTabInModel(config, SettingsUtils.getOrEmpty(settings, CFG_SETTINGS_TAB));
        loadAdvancedSettingsTabInModel(config, settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        loadEncryptionSettingsTabInModel(config, SettingsUtils.getOrEmpty(settings, CFG_ENCRYPTION_SETTINGS_TAB));
        config.setTableSpecConfig(loadTableSpecConfig(settings));
    }

    @Override
    public void saveInModel(final ExcelMultiTableReadConfig config, final NodeSettingsWO settings) {
        if (config.hasTableSpecConfig()) {
            m_tableSpecConfigSerializer.save(config.getTableSpecConfig(),
                settings.addNodeSettings(CFG_TABLE_SPEC_CONFIG));
        }
        saveSettingsTab(config, SettingsUtils.getOrAdd(settings, CFG_SETTINGS_TAB));
        saveAdvancedSettingsTab(config, settings.addNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        saveEncryptionSettingsTab(config, settings.addNodeSettings(CFG_ENCRYPTION_SETTINGS_TAB));
    }

    @Override
    public void validate(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings)
        throws InvalidSettingsException {
        validateSettingsTab(settings.getNodeSettings(CFG_SETTINGS_TAB));
        validateAdvancedSettingsTab(settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        validateEncryptionSettingsTab(config, SettingsUtils.getOrEmpty(settings, CFG_ENCRYPTION_SETTINGS_TAB));
    }

    @Override
    public void saveInDialog(final ExcelMultiTableReadConfig config, final NodeSettingsWO settings)
        throws InvalidSettingsException {
        saveInModel(config, settings);
    }

    private static void loadSettingsTabInDialog(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings) {
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
        excelConfig.setEmptyColHeaderPrefix(settings.getString(CFG_EMPTY_COL_HEADER_PREFIX, "empty_"));
        excelConfig.setColumnNameMode(
            ColumnNameMode.valueOf(settings.getString(CFG_COLUMN_NAME_MODE, ColumnNameMode.EXCEL_COL_NAME.name())));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER, true));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_NUMBER, 1) - 1);
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(false);
    }

    private static void loadSettingsTabInModel(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings)
        throws InvalidSettingsException {
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
        //Check for backwards compatability
        if (settings.containsKey(CFG_COLUMN_NAME_MODE)) {
            excelConfig.setColumnNameMode(ColumnNameMode.loadValueInModel(settings.getString(CFG_COLUMN_NAME_MODE)));
        }
        //Check for backwards compatability
        if (settings.containsKey(CFG_EMPTY_COL_HEADER_PREFIX)) {
            excelConfig.setEmptyColHeaderPrefix(settings.getString(CFG_EMPTY_COL_HEADER_PREFIX));
        }
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_NUMBER) - 1);
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(true);
        tableReadConfig.setUseRowIDIdx(excelConfig.getRowIdGeneration() == RowIDGeneration.COLUMN);
        tableReadConfig.setRowIDIdx(0);
        tableReadConfig.setDecorateRead(false);
    }

    private static void saveSettingsTab(final ExcelMultiTableReadConfig config, final NodeSettingsWO settings) {
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
        settings.addString(CFG_EMPTY_COL_HEADER_PREFIX, excelConfig.getEmptyColHeaderPrefix());
        settings.addString(CFG_COLUMN_NAME_MODE, excelConfig.getColumnNameMode().name());
    }

    static void validateSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
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
        //Check for backwards compatability
        if (settings.containsKey(CFG_COLUMN_NAME_MODE)) {
            settings.getString(CFG_COLUMN_NAME_MODE);
        }
        //Check for backwards compatability
        if (settings.containsKey(CFG_EMPTY_COL_HEADER_PREFIX)) {
            settings.getString(CFG_EMPTY_COL_HEADER_PREFIX);
        }
    }

    private static void loadAdvancedSettingsTabInDialog(final ExcelMultiTableReadConfig config,
        final NodeSettingsRO settings) {
        config.setFailOnDifferingSpecs(settings.getBoolean(CFG_FAIL_ON_DIFFERING_SPECS, true));
        // added in 4.4.0
        config.setAppendItemIdentifierColumn(settings.getBoolean(CFG_APPEND_PATH_COLUMN, false));
        config.setItemIdentifierColumnName(settings.getString(CFG_PATH_COLUMN_NAME, "Path"));
        config.setSkipEmptyColumns(settings.getBoolean(CFG_SKIP_EMPTY_COLS, false));
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setSkipEmptyRows(settings.getBoolean(CFG_SKIP_EMPTY_ROWS, true));
        tableReadConfig.setLimitRowsForSpec(settings.getBoolean(CFG_LIMIT_DATA_ROWS_SCANNED, true));
        tableReadConfig.setMaxRowsForSpec(
            settings.getLong(CFG_MAX_DATA_ROWS_SCANNED, AbstractTableReadConfig.DEFAULT_ROWS_FOR_SPEC_GUESSING));
        config.setSaveTableSpecConfig(settings.getBoolean(CFG_SAVE_TABLE_SPEC_CONFIG, true));

        // added in 5.2.2 (AP-19239); we don't want older workflows to make the check automatically
        config.setCheckSavedTableSpec(settings.getBoolean(CFG_CHECK_TABLE_SPEC, false));

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

    private static void loadAdvancedSettingsTabInModel(final ExcelMultiTableReadConfig config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        config.setFailOnDifferingSpecs(settings.getBoolean(CFG_FAIL_ON_DIFFERING_SPECS));
        // added in 4.4.0
        if (settings.containsKey(CFG_APPEND_PATH_COLUMN)) {
            config.setAppendItemIdentifierColumn(settings.getBoolean(CFG_APPEND_PATH_COLUMN));
            config.setItemIdentifierColumnName(settings.getString(CFG_PATH_COLUMN_NAME));
        }
        if (settings.containsKey(CFG_SKIP_EMPTY_COLS)) {
            config.setSkipEmptyColumns(settings.getBoolean(CFG_SKIP_EMPTY_COLS));
        }
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setSkipEmptyRows(settings.getBoolean(CFG_SKIP_EMPTY_ROWS));
        tableReadConfig.setLimitRowsForSpec(settings.getBoolean(CFG_LIMIT_DATA_ROWS_SCANNED));
        tableReadConfig.setMaxRowsForSpec(settings.getLong(CFG_MAX_DATA_ROWS_SCANNED));

        // added in 4.3.1
        if (settings.containsKey(CFG_SAVE_TABLE_SPEC_CONFIG)) {
            config.setSaveTableSpecConfig(settings.getBoolean(CFG_SAVE_TABLE_SPEC_CONFIG));
        }

        // added in 5.2.2 (AP-19239); we don't want existing workflows to make the check
        config.setCheckSavedTableSpec(settings.getBoolean(CFG_CHECK_TABLE_SPEC, false));

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

    private static void saveAdvancedSettingsTab(final ExcelMultiTableReadConfig config, final NodeSettingsWO settings) {
        settings.addBoolean(CFG_FAIL_ON_DIFFERING_SPECS, config.failOnDifferingSpecs());
        settings.addBoolean(CFG_APPEND_PATH_COLUMN, config.appendItemIdentifierColumn());
        settings.addString(CFG_PATH_COLUMN_NAME, config.getItemIdentifierColumnName());
        settings.addBoolean(CFG_SKIP_EMPTY_COLS, config.skipEmptyColumns());
        final DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        settings.addBoolean(CFG_SKIP_EMPTY_ROWS, tableReadConfig.skipEmptyRows());
        settings.addBoolean(CFG_LIMIT_DATA_ROWS_SCANNED, tableReadConfig.limitRowsForSpec());
        settings.addLong(CFG_MAX_DATA_ROWS_SCANNED, tableReadConfig.getMaxRowsForSpec());
        settings.addBoolean(CFG_SAVE_TABLE_SPEC_CONFIG, config.saveTableSpecConfig());
        settings.addBoolean(CFG_CHECK_TABLE_SPEC, config.checkSavedTableSpec());
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addBoolean(CFG_USE_15_DIGITS_PRECISION, excelConfig.isUse15DigitsPrecision());
        settings.addBoolean(CFG_SKIP_HIDDEN_COLS, excelConfig.isSkipHiddenCols());
        settings.addBoolean(CFG_SKIP_HIDDEN_ROWS, excelConfig.isSkipHiddenRows());
        settings.addBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS, excelConfig.isReplaceEmptyStringsWithMissings());
        settings.addBoolean(CFG_REEVALUATE_FORMULAS, excelConfig.isReevaluateFormulas());
        settings.addString(CFG_FORMULA_ERROR_HANDLING, excelConfig.getFormulaErrorHandling().name());
        settings.addString(CFG_FORMULA_ERROR_PATTERN, excelConfig.getErrorPattern());
    }

    static void validateAdvancedSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
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
        // added in 4.3.1
        if (settings.containsKey(CFG_SAVE_TABLE_SPEC_CONFIG)) {
            settings.getBoolean(CFG_SAVE_TABLE_SPEC_CONFIG);
        }

        // added in 5.2.2
        if (settings.containsKey(CFG_CHECK_TABLE_SPEC)) {
            settings.getBoolean(CFG_CHECK_TABLE_SPEC);
        }

        // added in 4.4.0
        if (settings.containsKey(CFG_APPEND_PATH_COLUMN)) {
            settings.getBoolean(CFG_APPEND_PATH_COLUMN);
            settings.getString(CFG_PATH_COLUMN_NAME);
        }
    }

    private static void loadEncryptionSettingsTabInModel(final ExcelMultiTableReadConfig config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        // this option has been added later with 4.4.0;
        // check the existence of the setting for sake of backwards compatibility
        if (settings.containsKey(ExcelTableReaderConfig.CFG_PASSWORD)) {
            excelConfig.getAuthenticationSettingsModel().loadSettingsFrom(settings);
        }
    }

    static NodeSettingsRO getOrDefaultEncryptionSettingsTab(final NodeSettingsRO settings) {
        if (settings.containsKey(CFG_ENCRYPTION_SETTINGS_TAB)) {
            try {
                return settings.getNodeSettings(CFG_ENCRYPTION_SETTINGS_TAB);
            } catch (InvalidSettingsException e) {
                LOGGER.error("Failed to fetch encryption settings despite the corresponding key being present.", e);
            }
        }
        // this tab was added in 4.4.0 therefore workflows with 4.3.0 don't have it
        // the SettingsModelAuthentication will fail to load in the dialog if its subsetting is missing
        // hence we create a dummy.
        final NodeSettings encryptionSettings = new NodeSettings(CFG_ENCRYPTION_SETTINGS_TAB);
        encryptionSettings.addNodeSettings(ExcelTableReaderConfig.CFG_PASSWORD);
        return encryptionSettings;
    }

    private static void saveEncryptionSettingsTab(final ExcelMultiTableReadConfig config,
        final NodeSettingsWO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.getAuthenticationSettingsModel().saveSettingsTo(settings);
    }

    static void validateEncryptionSettingsTab(final ExcelMultiTableReadConfig config, final NodeSettingsRO settings)
        throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        // this option has been added later; check the existence of the setting for sake of backwards compatibility
        if (settings.containsKey(ExcelTableReaderConfig.CFG_PASSWORD)) {
            excelConfig.getAuthenticationSettingsModel().validateSettings(settings);
        }
    }

    @Override
    public ConfigID createFromConfig(final ExcelMultiTableReadConfig config) {
        final NodeSettings settings = new NodeSettings("excel_reader");
        saveConfigIDSettingsTab(config, settings.addNodeSettings(CFG_SETTINGS_TAB));
        saveConfigIDAdvancedSettingsTab(config, settings.addNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        return new NodeSettingsConfigID(settings);
    }

    private static void saveConfigIDSettingsTab(final ExcelMultiTableReadConfig config, final NodeSettingsWO settings) {
        // currently all options of the settings tab affect the spec
        saveSettingsTab(config, settings);
    }

    private static void saveConfigIDAdvancedSettingsTab(final ExcelMultiTableReadConfig config,
        final NodeSettingsWO advancedSettings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        advancedSettings.addBoolean(CFG_SKIP_HIDDEN_COLS, excelConfig.isSkipHiddenCols());
        advancedSettings.addBoolean(CFG_SKIP_EMPTY_COLS, config.skipEmptyColumns());
        advancedSettings.addBoolean(CFG_SKIP_HIDDEN_ROWS, excelConfig.isSkipHiddenRows());
        advancedSettings.addString(CFG_FORMULA_ERROR_HANDLING, excelConfig.getFormulaErrorHandling().name());
        advancedSettings.addBoolean(CFG_REEVALUATE_FORMULAS, excelConfig.isReevaluateFormulas());
        advancedSettings.addBoolean(CFG_USE_15_DIGITS_PRECISION, excelConfig.isUse15DigitsPrecision());
        advancedSettings.addBoolean(CFG_REPLACE_EMPTY_STRINGS_WITH_MISSINGS,
            excelConfig.isReplaceEmptyStringsWithMissings());
        DefaultTableReadConfig<ExcelTableReaderConfig> tableReadConfig = config.getTableReadConfig();
        advancedSettings.addBoolean(CFG_LIMIT_DATA_ROWS_SCANNED, tableReadConfig.limitRowsForSpec());
        advancedSettings.addLong(CFG_MAX_DATA_ROWS_SCANNED, tableReadConfig.getMaxRowsForSpec());
    }

    @Override
    public ConfigID createFromSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        // Adjust this method if you add new settings that affect the spec
        return new NodeSettingsConfigID(settings.getNodeSettings("excel_reader"));
    }

}
