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
 * TODO implement once dialog is extended with more settings
 *
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

    private static final String CFG_SETTINGS_TAB = "settings";

    private static final String CGF_SHEET_SELECTION = "sheet_selection";

    private static final String CGF_SHEET_NAME = "sheet_name";

    private static final String CGF_SHEET_INDEX = "sheet_index";

    private static final String CFG_HAS_COLUMN_HEADER = "table_contains_column_names";

    private static final String CFG_COLUMN_HEADER_ROW_IDX = "column_names_row_number";

    private static final String CFG_ADVANCED_SETTINGS_TAB = "advanced_settings";

    private static final String CGF_USE_15_DIGITS_PRECISION = "use_15_digits_precision";

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
            SheetSelection.valueOf(settings.getString(CGF_SHEET_SELECTION, SheetSelection.FIRST.name())));
        excelConfig.setSheetName(settings.getString(CGF_SHEET_NAME, ""));
        excelConfig.setSheetIdx(settings.getInt(CGF_SHEET_INDEX, 1));
        final DefaultTableReadConfig<?> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(false);
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER, true));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_IDX, 1) - 1);
    }

    private static void loadSettingsTabInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setSheetSelection(SheetSelection.loadValueInModel(settings.getString(CGF_SHEET_SELECTION)));
        excelConfig.setSheetName(settings.getString(CGF_SHEET_NAME));
        excelConfig.setSheetIdx(settings.getInt(CGF_SHEET_INDEX));
        final DefaultTableReadConfig<?> tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setAllowShortRows(true);
        tableReadConfig.setLimitRowsForSpec(false);
        tableReadConfig.setUseColumnHeaderIdx(settings.getBoolean(CFG_HAS_COLUMN_HEADER));
        tableReadConfig.setColumnHeaderIdx(settings.getLong(CFG_COLUMN_HEADER_ROW_IDX) - 1);
    }

    private static void saveSettingsTab(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addString(CGF_SHEET_SELECTION, excelConfig.getSheetSelection().name());
        settings.addString(CGF_SHEET_NAME, excelConfig.getSheetName());
        settings.addInt(CGF_SHEET_INDEX, excelConfig.getSheetIdx());
        final DefaultTableReadConfig<?> tableReadConfig = config.getTableReadConfig();
        settings.addBoolean(CFG_HAS_COLUMN_HEADER, tableReadConfig.useColumnHeaderIdx());
        settings.addLong(CFG_COLUMN_HEADER_ROW_IDX, tableReadConfig.getColumnHeaderIdx() + 1);
    }

    public static void validateSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
        settings.getString(CGF_SHEET_SELECTION);
        settings.getString(CGF_SHEET_NAME);
        settings.getInt(CGF_SHEET_INDEX);
        settings.getBoolean(CFG_HAS_COLUMN_HEADER);
        settings.getLong(CFG_COLUMN_HEADER_ROW_IDX);
    }

    private static void loadAdvancedSettingsTabInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CGF_USE_15_DIGITS_PRECISION, true));
    }

    private static void loadAdvancedSettingsTabInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CGF_USE_15_DIGITS_PRECISION));
    }

    private static void saveAdvancedSettingsTab(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addBoolean(CGF_USE_15_DIGITS_PRECISION, excelConfig.isUse15DigitsPrecision());
    }

    public static void validateAdvancedSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
        settings.getBoolean(CGF_USE_15_DIGITS_PRECISION);
    }

}
