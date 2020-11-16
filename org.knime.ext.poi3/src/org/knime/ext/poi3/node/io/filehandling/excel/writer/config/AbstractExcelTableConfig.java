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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.config;

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
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.filehandling.core.defaultnodesettings.filechooser.AbstractSettingsModelFileChooser;

/**
 * Abstract implementation of the {@link ExcelTableConfig}. This class contains all settings that the writer and
 * appender node have in common.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 * @param <T> an instance of {@link AbstractSettingsModelFileChooser}
 */
public abstract class AbstractExcelTableConfig<T extends AbstractSettingsModelFileChooser<?>>
    implements ExcelTableConfig {

    /** Settings key of the file chooser. */
    protected static final String CFG_FILE_CHOOSER = "file_selection";

    private static final String DEFAULT_SHEET_NAME_PREFIX = "default_";

    private static final String CFG_MISSING_VALUE_PATTERN = "missing_value_pattern";

    private static final String CFG_WRITE_ROW_KEY = "write_row_key";

    private static final String CFG_WRITE_COLUMN_HEADER = "write_column_header";

    private static final String CFG_REPLACE_MISSINGS = "replace_missings";

    private static final String CFG_LANDSCAPE = "layout";

    private static final String CFG_AUTOSIZE = "autosize_columns";

    private static final String CFG_PAPER_SIZE = "paper_size";

    private static final String CFG_SHEET_NAMES = "sheet_names";

    private final T m_fileChooser;

    private final SettingsModelBoolean m_replaceMissings;

    private final SettingsModelString m_missingValPattern;

    private final SettingsModelBoolean m_writeRowKey;

    private final SettingsModelBoolean m_writeColHeader;

    private final SettingsModelBoolean m_autoSize;

    private final SettingsModelString m_paperSize;

    private final SettingsModelString m_landscape;

    private String[] m_sheetNames;

    /**
     * Constructor.
     *
     * @param portsCfg the {@link PortsConfiguration}
     * @param fileChooser instance of {@link AbstractSettingsModelFileChooser}
     * @param sheetGrpID the ID of the sheet group
     */
    protected AbstractExcelTableConfig(final PortsConfiguration portsCfg, final T fileChooser,
        final String sheetGrpID) {
        m_fileChooser = fileChooser;
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
        m_sheetNames = IntStream.range(0, portsCfg.getInputPortLocation().get(sheetGrpID).length)//
            .mapToObj(AbstractExcelTableConfig::createDefaultSheetName)//
            .toArray(String[]::new);
    }

    private static String createDefaultSheetName(final int idx) {
        return DEFAULT_SHEET_NAME_PREFIX + (idx + 1);
    }

    void setSheetNames(final String[] sheetNames) {
        m_sheetNames = sheetNames;
    }

    /**
     * Returns the concrete instance of {@link AbstractSettingsModelFileChooser}.
     *
     * @return the concrete instance of {@link AbstractSettingsModelFileChooser}
     */
    public final T getFileChooserModel() {
        return m_fileChooser;
    }

    @Override
    public final String[] getSheetNames() {
        return m_sheetNames;
    }

    /**
     * Returns the write column header settings model.
     *
     * @return the write column header settings model
     */
    public final SettingsModelBoolean getWriteColHeaderModel() {
        return m_writeColHeader;
    }

    /**
     * Returns the write row key settings model.
     *
     * @return the write row key settings model
     */
    public final SettingsModelBoolean getWriteRowKeyModel() {
        return m_writeRowKey;
    }

    /**
     * Returns the replace missings model
     *
     * @return the replace missings model
     */
    public final SettingsModelBoolean getReplaceMissingsModel() {
        return m_replaceMissings;
    }

    /**
     * Returns the missing value pattern model.
     *
     * @return the missing value pattern model
     */
    public final SettingsModelString getMissingValPatternModel() {
        return m_missingValPattern;
    }

    /**
     * Returns the auto size model.
     *
     * @return the auto size model
     */
    public final SettingsModelBoolean getAutoSizeModel() {
        return m_autoSize;
    }

    /**
     * Returns the landscape model.
     *
     * @return the landscape model
     */
    public final SettingsModelString getLandscapeModel() {
        return m_landscape;
    }

    /**
     * Returns the paper size model.
     *
     * @return the paper size model
     */
    public final SettingsModelString getPaperSizeModel() {
        return m_paperSize;
    }

    @Override
    public final Optional<String> getMissingValPattern() {
        if (m_replaceMissings.getBooleanValue()) {
            return Optional.of(m_missingValPattern.getStringValue());
        } else {
            return Optional.empty();
        }
    }

    @Override
    public final boolean useAutoSize() {
        return m_autoSize.getBooleanValue();
    }

    @Override
    public final boolean useLandscape() {
        return Orientation.valueOf(m_landscape.getStringValue()).useLandscape();
    }

    @Override
    public final short getPaperSize() {
        return PaperSize.valueOf(m_paperSize.getStringValue()).getPrintSetup();
    }

    @Override
    public final boolean writeRowKey() {
        return m_writeRowKey.getBooleanValue();
    }

    @Override
    public final boolean writeColHeaders() {
        return m_writeColHeader.getBooleanValue();
    }

    /**
     * Validates the {@link NodeSettingsRO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsRO} to validate
     * @throws InvalidSettingsException - If the validation failed
     */
    public final void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_fileChooser.validateSettings(settings);
        m_writeRowKey.validateSettings(settings);
        m_writeColHeader.validateSettings(settings);
        m_replaceMissings.validateSettings(settings);
        m_missingValPattern.validateSettings(settings);
        m_autoSize.validateSettings(settings);
        m_landscape.validateSettings(settings);
        m_paperSize.validateSettings(settings);
        validateSheets(settings);
        validateAdditionalSettingsForModel(settings);
    }

    /**
     * Validates any additional settings that are specific to the extending class.
     *
     * @param settings the {@link NodeSettingsRO} to validated
     * @throws InvalidSettingsException - If the settings are invalid
     */
    protected abstract void validateAdditionalSettingsForModel(final NodeSettingsRO settings)
        throws InvalidSettingsException;

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
    public final void loadSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_fileChooser.loadSettingsFrom(settings);
        loadSheetNamesForModel(settings);
        m_writeRowKey.loadSettingsFrom(settings);
        m_writeColHeader.loadSettingsFrom(settings);
        m_replaceMissings.loadSettingsFrom(settings);
        m_missingValPattern.loadSettingsFrom(settings);
        m_autoSize.loadSettingsFrom(settings);
        m_landscape.loadSettingsFrom(settings);
        m_paperSize.loadSettingsFrom(settings);
        loadAdditionalSettingsForModel(settings);
    }

    /**
     * Loads any additional settings that are specific to the extending class.
     *
     * @param settings the {@link NodeSettingsRO} to read from
     * @throws InvalidSettingsException - If the settings could not be loaded
     */
    protected abstract void loadAdditionalSettingsForModel(final NodeSettingsRO settings)
        throws InvalidSettingsException;

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
    public final void saveSettingsForModel(final NodeSettingsWO settings) {
        m_fileChooser.saveSettingsTo(settings);
        saveSheetNames(settings, m_sheetNames);
        m_writeRowKey.saveSettingsTo(settings);
        m_writeColHeader.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_autoSize.saveSettingsTo(settings);
        m_landscape.saveSettingsTo(settings);
        m_paperSize.saveSettingsTo(settings);
        saveAdditionalSettingsForModel(settings);
    }

    /**
     * Saves any additional settings that are specific to the extending class.
     *
     * @param settings the {@link NodeSettingsWO} to write to
     */
    protected abstract void saveAdditionalSettingsForModel(final NodeSettingsWO settings);

    /**
     * Saves the sheet names.
     *
     * @param settings the {@link NodeSettingsWO} to write the sheet names to
     * @param array the array to write
     */
    public static void saveSheetNames(final NodeSettingsWO settings, final String[] array) {
        settings.addStringArray(CFG_SHEET_NAMES, array);
    }

    /**
     * Loads the sheet names in the {@link NodeDialog}.
     *
     * @param settings the {@link NodeSettingsRO} containing the sheet names
     */
    public final void loadSheetsInDialog(final NodeSettingsRO settings) {
        try {
            loadSheetNamesForModel(settings);
        } catch (final InvalidSettingsException e) { // NOSONAR we want to use defaults
            final String[] defaultSheets = IntStream.range(0, m_sheetNames.length)//
                .mapToObj(AbstractExcelTableConfig::createDefaultSheetName)//
                .toArray(String[]::new);
            setSheetNames(settings.getStringArray(CFG_SHEET_NAMES, defaultSheets));
        }
    }

}
