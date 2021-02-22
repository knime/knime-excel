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
 *   Nov 16, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import java.util.HashSet;
import java.util.Optional;
import java.util.Set;
import java.util.stream.IntStream;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialog;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModel;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetNameExistsHandling;
import org.knime.filehandling.core.defaultnodesettings.EnumConfig;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

/**
 * {@link ExcelTableConfig} of the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterConfig implements ExcelTableConfig {

    private static final ExcelFormat DEFAULT_EXCEL_FORMAT = ExcelFormat.XLSX;

    private static final SheetNameExistsHandling DEFAULT_SHEET_EXISTS_HANDLING = SheetNameExistsHandling.FAIL;

    private static final String CFG_EXCEL_FORMAT = "excel_format";

    private static final String CFG_FILE_CHOOSER = "file_selection";

    private static final String DEFAULT_SHEET_NAME_PREFIX = "default_";

    private static final String CFG_SHEET_NAMES = "sheet_names";

    private static final String CFG_SHEET_EXISTS = "if_sheet_exists";

    private static final String CFG_WRITE_ROW_KEY = "write_row_key";

    private static final String CFG_WRITE_COLUMN_HEADER = "write_column_header";

    private static final String CFG_MISSING_VALUE_PATTERN = "missing_value_pattern";

    private static final String CFG_REPLACE_MISSINGS = "replace_missings";

    private static final String CFG_LANDSCAPE = "layout";

    private static final String CFG_AUTOSIZE = "autosize_columns";

    private static final String CFG_PAPER_SIZE = "paper_size";

    private static final String CFG_EVALUATE_FORMULAS = "evaluate_formulas";

    private static final String CFG_OPEN_FILE_AFTER_EXEC = "open_file_after_exec" + SettingsModel.CFGKEY_INTERNAL;

    private final SettingsModelString m_excelFormat;

    private final SettingsModelWriterFileChooser m_fileChooser;

    private String[] m_sheetNames;

    private final SettingsModelString m_sheetExistsHandling;

    private final SettingsModelBoolean m_replaceMissings;

    private final SettingsModelString m_missingValPattern;

    private final SettingsModelBoolean m_writeRowKey;

    private final SettingsModelBoolean m_writeColHeader;

    private final SettingsModelBoolean m_evaluateFormulas;

    private final SettingsModelBoolean m_autoSize;

    private final SettingsModelString m_paperSize;

    private final SettingsModelString m_landscape;

    private final SettingsModelBoolean m_openFileAfterExec;

    /**
     * Constructor.
     *
     * @param portsCfg the {@link PortsConfiguration}
     */
    ExcelTableWriterConfig(final PortsConfiguration portsCfg) {
        m_excelFormat = new SettingsModelString(CFG_EXCEL_FORMAT, DEFAULT_EXCEL_FORMAT.name());
        m_fileChooser = new SettingsModelWriterFileChooser(CFG_FILE_CHOOSER, portsCfg,
            ExcelTableWriterNodeFactory.FS_CONNECT_GRP_ID, EnumConfig.create(FilterMode.FILE),
            EnumConfig.create(FileOverwritePolicy.FAIL, FileOverwritePolicy.OVERWRITE, FileOverwritePolicy.APPEND),
            DEFAULT_EXCEL_FORMAT.getFileExtension());
        m_sheetNames =
            IntStream.range(0, portsCfg.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length)//
                .mapToObj(ExcelTableWriterConfig::createDefaultSheetName)//
                .toArray(String[]::new);
        m_sheetExistsHandling = new SettingsModelString(CFG_SHEET_EXISTS, DEFAULT_SHEET_EXISTS_HANDLING.name()) {

            @Override
            protected void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
                super.validateSettingsForModel(settings);
                final String setting = settings.getString(getConfigName());
                try {
                    SheetNameExistsHandling.valueOf(setting);
                } catch (final IllegalArgumentException e) {
                    throw new InvalidSettingsException(
                        String.format("No if sheet exists option associated with '%s'", setting), e);
                }
            }
        };
        m_missingValPattern = new SettingsModelString(CFG_MISSING_VALUE_PATTERN, "");
        m_writeRowKey = new SettingsModelBoolean(CFG_WRITE_ROW_KEY, false);
        m_writeColHeader = new SettingsModelBoolean(CFG_WRITE_COLUMN_HEADER, true);
        m_evaluateFormulas = new SettingsModelBoolean(CFG_EVALUATE_FORMULAS, false);
        m_replaceMissings = new SettingsModelBoolean(CFG_REPLACE_MISSINGS, false);
        m_autoSize = new SettingsModelBoolean(CFG_AUTOSIZE, false);
        m_landscape = new SettingsModelString(CFG_LANDSCAPE, Orientation.PORTRAIT.name()) {

            @Override
            protected void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
                super.validateSettingsForModel(settings);
                final String setting = settings.getString(getConfigName());
                try {
                    Orientation.valueOf(setting);
                } catch (final IllegalArgumentException e) {
                    throw new InvalidSettingsException(String.format("No orientation associated with '%s'", setting),
                        e);
                }
            }
        };
        m_paperSize = new SettingsModelString(CFG_PAPER_SIZE, ExcelConstants.DEFAULT_PAPER_SIZE.name());

        // introduced with 4.3.1 so we need to load it backwards compatible
        m_openFileAfterExec = new SettingsModelBoolean(CFG_OPEN_FILE_AFTER_EXEC, false) {

            @Override
            protected void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
                if (settings.containsKey(getConfigName())) {
                    super.validateSettingsForModel(settings);
                }
            }

            @Override
            protected void loadSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
                if (settings.containsKey(getConfigName())) {
                    super.loadSettingsForModel(settings);
                }
            }
        };
    }

    private static String createDefaultSheetName(final int idx) {
        return DEFAULT_SHEET_NAME_PREFIX + (idx + 1);
    }

    void setSheetNames(final String[] sheetNames) {
        m_sheetNames = sheetNames;
    }

    /**
     * Returns the {@link SettingsModelString} storing the {@link ExcelFormat}.
     *
     * @return the {@link SettingsModelString} storing the {@link ExcelFormat}
     */
    SettingsModelString getExcelFormatModel() {
        return m_excelFormat;
    }

    /**
     * Returns the {@link SettingsModelWriterFileChooser}.
     *
     * @return the {@link SettingsModelWriterFileChooser}
     */
    SettingsModelWriterFileChooser getFileChooserModel() {
        return m_fileChooser;
    }

    @Override
    public String[] getSheetNames() {
        return m_sheetNames;
    }

    /**
     * Returns the sheet exists handling model.
     *
     * @return the sheet exists handling model
     */
    SettingsModelString getSheetExistsHandlingModel() {
        return m_sheetExistsHandling;
    }

    /**
     * Returns the write column header settings model.
     *
     * @return the write column header settings model
     */
    SettingsModelBoolean getWriteColHeaderModel() {
        return m_writeColHeader;
    }

    /**
     * Returns the write row key settings model.
     *
     * @return the write row key settings model
     */
    SettingsModelBoolean getWriteRowKeyModel() {
        return m_writeRowKey;
    }

    /**
     * Returns the replace missings model
     *
     * @return the replace missings model
     */
    SettingsModelBoolean getReplaceMissingsModel() {
        return m_replaceMissings;
    }

    /**
     * Returns the missing value pattern model.
     *
     * @return the missing value pattern model
     */
    SettingsModelString getMissingValPatternModel() {
        return m_missingValPattern;
    }

    /**
     * Returns the evaluate formulas model.
     *
     * @return the evaluate formulas model
     */
    SettingsModelBoolean getEvaluateFormulasModel() {
        return m_evaluateFormulas;
    }

    /**
     * Returns the auto size model.
     *
     * @return the auto size model
     */
    SettingsModelBoolean getAutoSizeModel() {
        return m_autoSize;
    }

    /**
     * Returns the landscape model.
     *
     * @return the landscape model
     */
    SettingsModelString getLandscapeModel() {
        return m_landscape;
    }

    /**
     * Returns the paper size model.
     *
     * @return the paper size model
     */
    SettingsModelString getPaperSizeModel() {
        return m_paperSize;
    }

    /**
     * Returns the open file after execution model.
     *
     * @return the open file after execution model
     */
    SettingsModelBoolean getOpenFileAfterExecModel() {
        return m_openFileAfterExec;
    }

    @Override
    public ExcelFormat getExcelFormat() {
        return ExcelFormat.valueOf(m_excelFormat.getStringValue());
    }

    @Override
    public Optional<String> getMissingValPattern() {
        if (m_replaceMissings.getBooleanValue()) {
            return Optional.of(m_missingValPattern.getStringValue());
        } else {
            return Optional.empty();
        }
    }

    @Override
    public boolean useAutoSize() {
        return m_autoSize.getBooleanValue();
    }

    @Override
    public boolean useLandscape() {
        return Orientation.valueOf(m_landscape.getStringValue()).useLandscape();
    }

    @Override
    public short getPaperSize() {
        return PaperSize.valueOf(m_paperSize.getStringValue()).getPrintSetup();
    }

    @Override
    public boolean writeRowKey() {
        return m_writeRowKey.getBooleanValue();
    }

    @Override
    public boolean writeColHeaders() {
        return m_writeColHeader.getBooleanValue();
    }

    @Override
    public boolean evaluate() {
        return m_evaluateFormulas.getBooleanValue();
    }

    @Override
    public boolean abortIfSheetExists() {
        return SheetNameExistsHandling.valueOf(m_sheetExistsHandling.getStringValue()) == SheetNameExistsHandling.FAIL;
    }

    /**
     * Validates the {@link NodeSettingsRO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsRO} to validate
     * @throws InvalidSettingsException - If the validation failed
     */
    void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_excelFormat.validateSettings(settings);
        m_fileChooser.validateSettings(settings);
        validateSheets(settings);
        m_sheetExistsHandling.validateSettings(settings);
        m_writeRowKey.validateSettings(settings);
        m_writeColHeader.validateSettings(settings);
        m_replaceMissings.validateSettings(settings);
        m_missingValPattern.validateSettings(settings);
        m_evaluateFormulas.validateSettings(settings);
        m_autoSize.validateSettings(settings);
        m_landscape.validateSettings(settings);
        m_paperSize.validateSettings(settings);
        m_openFileAfterExec.validateSettings(settings);
    }

    private static void validateSheets(final NodeSettingsRO settings) throws InvalidSettingsException {
        final String[] sheetNames = settings.getStringArray(CFG_SHEET_NAMES);
        final Set<String> uniqueNames = new HashSet<>();
        for (final String sheetName : sheetNames) {
            validateSheetName(sheetName);
            if (!uniqueNames.add(sheetName)) {
                throw new InvalidSettingsException("The sheet names must be unique.");
            }
        }
    }

    private static void validateSheetName(final String sheetName) throws InvalidSettingsException {
        if (sheetName == null || sheetName.trim().isEmpty()) {
            throw new InvalidSettingsException("The sheet name may not be empty.");
        }
        if (sheetName.length() > ExcelConstants.MAX_SHEET_NAME_LENGTH) {
            throw new InvalidSettingsException("The sheet name may not contain more than 31 charcters.");
        }
        if (ExcelConstants.ILLEGAL_SHEET_NAME_PATTERN.matcher(sheetName).find()) {
            throw new InvalidSettingsException("The sheet name may not contain :, [, ], /, \\, * or ?.");
        }
    }

    /**
     * Loads the {@link NodeSettingsRO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsRO} to be loaded
     * @throws InvalidSettingsException - If loading the settings failed
     */
    void loadSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_excelFormat.loadSettingsFrom(settings);
        m_fileChooser.loadSettingsFrom(settings);
        loadSheetNamesForModel(settings);
        m_sheetExistsHandling.loadSettingsFrom(settings);
        m_writeRowKey.loadSettingsFrom(settings);
        m_writeColHeader.loadSettingsFrom(settings);
        m_replaceMissings.loadSettingsFrom(settings);
        m_missingValPattern.loadSettingsFrom(settings);
        m_evaluateFormulas.loadSettingsFrom(settings);
        m_autoSize.loadSettingsFrom(settings);
        m_landscape.loadSettingsFrom(settings);
        m_paperSize.loadSettingsFrom(settings);
        m_openFileAfterExec.loadSettingsFrom(settings);
    }

    private void loadSheetNamesForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        final String[] sheetNames = settings.getStringArray(CFG_SHEET_NAMES);
        IntStream.range(0, Math.min(m_sheetNames.length, sheetNames.length))//
            .forEach(idx -> m_sheetNames[idx] = sheetNames[idx]);
    }

    /**
     * Saves the {@link NodeSettingsWO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsWO} to save to
     */
    void saveSettingsForModel(final NodeSettingsWO settings) {
        m_excelFormat.saveSettingsTo(settings);
        m_fileChooser.saveSettingsTo(settings);
        saveSheetNames(settings, m_sheetNames);
        m_sheetExistsHandling.saveSettingsTo(settings);
        m_writeRowKey.saveSettingsTo(settings);
        m_writeColHeader.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
        m_autoSize.saveSettingsTo(settings);
        m_landscape.saveSettingsTo(settings);
        m_paperSize.saveSettingsTo(settings);
        m_openFileAfterExec.saveSettingsTo(settings);
    }

    /**
     * Saves the sheet names.
     *
     * @param settings the {@link NodeSettingsWO} to write the sheet names to
     * @param array the array to write
     */
    static void saveSheetNames(final NodeSettingsWO settings, final String[] array) {
        settings.addStringArray(CFG_SHEET_NAMES, array);
    }

    /**
     * Loads the sheet names in the {@link NodeDialog}.
     *
     * @param settings the {@link NodeSettingsRO} containing the sheet names
     */
    void loadSheetsInDialog(final NodeSettingsRO settings) {
        try {
            loadSheetNamesForModel(settings);
        } catch (final InvalidSettingsException e) { // NOSONAR we want to use defaults
            final String[] defaultSheets = IntStream.range(0, m_sheetNames.length)//
                .mapToObj(ExcelTableWriterConfig::createDefaultSheetName)//
                .toArray(String[]::new);
            setSheetNames(settings.getStringArray(CFG_SHEET_NAMES, defaultSheets));
        }
    }

}
