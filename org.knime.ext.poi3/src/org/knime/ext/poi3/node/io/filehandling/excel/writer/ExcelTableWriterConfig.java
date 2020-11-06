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
 *   Nov 9, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import java.util.EnumSet;
import java.util.HashSet;
import java.util.Optional;
import java.util.Set;
import java.util.stream.IntStream;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

/**
 * Stores the configuration for the {@link ExcelTableWriterNodeModel}.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
@SuppressWarnings("javadoc")
public final class ExcelTableWriterConfig {

    private static final String DEFAULT_SHEET_NAME_PREFIX = "default_";

    private static final ExcelFormat DEFAULT_EXCEL_FORMAT = ExcelFormat.XLSX;

    private static final String CFG_FILE_CHOOSER = "file_location";

    private static final String CFG_EXCEL_FORMAT = "excel_format";

    private static final String CFG_MISSING_VALUE_PATTERN = "missing_value_pattern";

    private static final String CFG_WRITE_ROW_KEY = "write_row_key";

    private static final String CFG_WRITE_COLUMN_HEADER = "write_column_header";

    private static final String CFG_REPLACE_MISSINGS = "replace_missings";

    private static final String CFG_LANDSCAPE = "layout";

    private static final String CFG_AUTOSIZE = "autosize_columns";

    private static final String CFG_PAPER_SIZE = "paper_size";

    private static final String CFG_SHEET_NAMES = "sheet_names";

    private final SettingsModelWriterFileChooser m_fileChooser;

    private final SettingsModelString m_excelFormat;

    private final SettingsModelBoolean m_replaceMissings;

    private final SettingsModelString m_missingValPattern;

    private final SettingsModelBoolean m_writeRowKey;

    private final SettingsModelBoolean m_writeColHeader;

    private final SettingsModelBoolean m_autoSize;

    private final SettingsModelString m_paperSize;

    private final SettingsModelString m_landscape;

    private String[] m_sheetNames;

    ExcelTableWriterConfig(final PortsConfiguration portsCfg) {
        m_fileChooser = new SettingsModelWriterFileChooser(CFG_FILE_CHOOSER, portsCfg,
            ExcelTableWriterNodeFactory.FS_CONNECT_GRP_ID, FilterMode.FILE, FileOverwritePolicy.FAIL,
            EnumSet.of(FileOverwritePolicy.FAIL, FileOverwritePolicy.OVERWRITE),
            DEFAULT_EXCEL_FORMAT.getFileExtension());
        m_excelFormat = new SettingsModelString(CFG_EXCEL_FORMAT, DEFAULT_EXCEL_FORMAT.name());
        m_missingValPattern = new SettingsModelString(CFG_MISSING_VALUE_PATTERN, "");
        m_writeRowKey = new SettingsModelBoolean(CFG_WRITE_ROW_KEY, false);
        m_writeColHeader = new SettingsModelBoolean(CFG_WRITE_COLUMN_HEADER, true);
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
                    throw new InvalidSettingsException(String.format("No orientation associated with %s", setting));
                }
            }
        };
        m_paperSize = new SettingsModelString(CFG_PAPER_SIZE, ExcelConstants.DEFAULT_PAPER_SIZE.name());
        m_sheetNames =
            IntStream.range(0, portsCfg.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length)//
                .mapToObj(ExcelTableWriterConfig::createDefaultSheetName)//
                .toArray(String[]::new);
    }

    private static String createDefaultSheetName(final int idx) {
        return DEFAULT_SHEET_NAME_PREFIX + (idx + 1);
    }

    void setSheetNames(final String[] sheetNames) {
        m_sheetNames = sheetNames;
    }

    SettingsModelString getExcelFormatModel() {
        return m_excelFormat;
    }

    SettingsModelWriterFileChooser getFileChooserModel() {
        return m_fileChooser;
    }

    String[] getSheetNames() {
        return m_sheetNames;
    }

    SettingsModelBoolean getWriteColHeaderModel() {
        return m_writeColHeader;
    }

    SettingsModelBoolean getWriteRowKeyModel() {
        return m_writeRowKey;
    }

    SettingsModelBoolean getReplaceMissingsModel() {
        return m_replaceMissings;
    }

    SettingsModelBoolean getAutoSizeModel() {
        return m_autoSize;
    }

    SettingsModelString getLandscapeModel() {
        return m_landscape;
    }

    SettingsModelString getPaperSizeModel() {
        return m_paperSize;
    }

    SettingsModelString getMissingValPatternModel() {
        return m_missingValPattern;
    }

    /**
     * Returns the selected {@link ExcelFormat}
     *
     * @return the selected {@link ExcelFormat}
     */
    public ExcelFormat getExcelFormat() {
        return ExcelFormat.valueOf(m_excelFormat.getStringValue());
    }

    /**
     * Returns the missing value pattern if selected.
     *
     * @return the missing value pattern if selected
     */
    public Optional<String> getMissingValPattern() {
        if (m_replaceMissings.getBooleanValue()) {
            return Optional.of(m_missingValPattern.getStringValue());
        } else {
            return Optional.empty();
        }
    }

    /**
     * Flag indicating whether or not the sheet's columns have to be auto sized.
     *
     * @return {@code true} if the sheet's columns have to be auto sized and {@code false} otherwise
     */
    public boolean useAutoSize() {
        return m_autoSize.getBooleanValue();
    }

    /**
     * Flag indicating whether or not the sheet uses the landscape layout.
     *
     * @return {@code true} if the sheet's layout has to be set to landscape and {@code false} otherwise
     */
    public boolean useLandscape() {
        return Orientation.valueOf(m_landscape.getStringValue()).useLandscape();
    }

    /**
     * Returns the paper size.
     *
     * @return the paper size
     */
    public short getPaperSize() {
        return PaperSize.valueOf(m_paperSize.getStringValue()).getPrintSetup();
    }

    /**
     * Flag indicating whether or not row keys must be written to the sheet.
     *
     * @return {@code true} if the row keys must be written to the sheet and {@code false} otherwise
     */
    public boolean writeRowKey() {
        return m_writeRowKey.getBooleanValue();
    }

    /**
     * Flag indicating whether or not column headers must be written to the sheet.
     *
     * @return {@code true} if the column headers must be written to the sheet and {@code false} otherwise
     */
    public boolean writeColHeaders() {
        return m_writeColHeader.getBooleanValue();
    }

    void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_fileChooser.validateSettings(settings);
        m_excelFormat.validateSettings(settings);
        m_writeRowKey.validateSettings(settings);
        m_writeColHeader.validateSettings(settings);
        m_replaceMissings.validateSettings(settings);
        m_missingValPattern.validateSettings(settings);
        m_autoSize.validateSettings(settings);
        m_landscape.validateSettings(settings);
        m_paperSize.validateSettings(settings);
        validateSheets(settings);
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

    void loadSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_excelFormat.loadSettingsFrom(settings);
        m_fileChooser.loadSettingsFrom(settings);
        loadSheetNamesForModel(settings);
        m_writeRowKey.loadSettingsFrom(settings);
        m_writeColHeader.loadSettingsFrom(settings);
        m_replaceMissings.loadSettingsFrom(settings);
        m_missingValPattern.loadSettingsFrom(settings);
        m_autoSize.loadSettingsFrom(settings);
        m_landscape.loadSettingsFrom(settings);
        m_paperSize.loadSettingsFrom(settings);
    }

    private void loadSheetNamesForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        final String[] sheetNames = settings.getStringArray(CFG_SHEET_NAMES);
        IntStream.range(0, Math.min(m_sheetNames.length, sheetNames.length))//
            .forEach(idx -> m_sheetNames[idx] = sheetNames[idx]);
    }

    void saveSettingsForModel(final NodeSettingsWO settings) {
        m_excelFormat.saveSettingsTo(settings);
        m_fileChooser.saveSettingsTo(settings);
        saveSheetNames(settings, m_sheetNames);
        m_writeRowKey.saveSettingsTo(settings);
        m_writeColHeader.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_autoSize.saveSettingsTo(settings);
        m_landscape.saveSettingsTo(settings);
        m_paperSize.saveSettingsTo(settings);
    }

    static void saveSheetNames(final NodeSettingsWO settings, final String[] array) {
        settings.addStringArray(CFG_SHEET_NAMES, array);
    }

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