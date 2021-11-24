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
package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

import java.util.Arrays;
import java.util.Optional;
import java.util.stream.Stream;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialog;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.filehandling.core.defaultnodesettings.EnumConfig;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

/**
 * {@link ExcelTableConfig} of the 'Excel Cell Updater' node.
 *
 * @author Moditha Hewasinghage,, KNIME GmbH, Berlin, Germany
 * @author Jannik Löscher, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelCellUpdaterConfig implements ExcelTableConfig {

    private static final ExcelFormat DEFAULT_EXCEL_FORMAT = ExcelFormat.XLSX;

    private static final String CFG_SRC_FILE_CHOOSER = "src_file_selection";

    private static final String CFG_CREATE_NEW_FILE = "create_new_file";

    private static final String CFG_DEST_FILE_CHOOSER = "dest_file_selection";

    private static final String CFG_SHEET_NAMES = "sheet_names";

    private static final String CFG_COORDINATE_COLUMN_NAMES = "coordinate_column_names";

    private static final String CFG_MISSING_VALUE_PATTERN = "missing_value_pattern";

    private static final String CFG_REPLACE_MISSINGS = "replace_missings";

    private static final String CFG_EVALUATE_FORMULAS = "evaluate_formulas";

    private final SettingsModelReaderFileChooser m_srcFileChooser;

    private final SettingsModelWriterFileChooser m_destFileChooser;

    private final SettingsModelBoolean m_createNewFile;

    private final SettingsModelBoolean m_evaluateFormulas;

    private final SettingsModelBoolean m_replaceMissings;

    private final SettingsModelString m_missingValPattern;

    private String[] m_sheetNames;

    private String[] m_coordinateColumnNames;

    /**
     * Constructor.
     *
     * @param portsCfg the {@link PortsConfiguration}
     */
    ExcelCellUpdaterConfig(final PortsConfiguration portsCfg) {
        m_srcFileChooser = new SettingsModelReaderFileChooser(CFG_SRC_FILE_CHOOSER, portsCfg,
            ExcelCellUpdaterNodeFactory.FS_CONNECT_GRP_ID, EnumConfig.create(FilterMode.FILE));

        m_destFileChooser = new SettingsModelWriterFileChooser(CFG_DEST_FILE_CHOOSER, portsCfg,
            ExcelCellUpdaterNodeFactory.FS_CONNECT_GRP_ID, EnumConfig.create(FilterMode.FILE),
            EnumConfig.create(FileOverwritePolicy.FAIL, FileOverwritePolicy.OVERWRITE),
            DEFAULT_EXCEL_FORMAT.getFileExtension());

        m_createNewFile = new SettingsModelBoolean(CFG_CREATE_NEW_FILE, false);
        m_evaluateFormulas = new SettingsModelBoolean(CFG_EVALUATE_FORMULAS, false);
        m_missingValPattern = new SettingsModelString(CFG_MISSING_VALUE_PATTERN, "");
        m_replaceMissings = new SettingsModelBoolean(CFG_REPLACE_MISSINGS, false);
        m_sheetNames = Stream.generate(() -> "")//
            .limit(portsCfg.getInputPortLocation().get(ExcelCellUpdaterNodeFactory.SHEET_GRP_ID).length)//
            .toArray(String[]::new);
        m_coordinateColumnNames = Arrays.copyOf(m_sheetNames, m_sheetNames.length);
    }

    /**
     * Validates the {@link NodeSettingsRO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsRO} to validate
     * @throws InvalidSettingsException - If the validation failed
     */
    void validateSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        // we need to create clones so that we can easily read the values saved in the settings to be able to perform further checks
        final SettingsModelReaderFileChooser srcClone = m_srcFileChooser.createCloneWithValidatedValue(settings);
        final SettingsModelBoolean createFileClone = m_createNewFile.createCloneWithValidatedValue(settings);
        if (createFileClone.getBooleanValue()) {
            final SettingsModelWriterFileChooser destClone = m_destFileChooser.createCloneWithValidatedValue(settings);
            var srcPath = srcClone.getLocation().getPath();
            var destPath = destClone.getLocation().getPath();
            var srcExtension = Optional.ofNullable(ExcelCellUpdaterNodeModel.getExcelFormat(srcPath));
            var destExtension = Optional.ofNullable(ExcelCellUpdaterNodeModel.getExcelFormat(destPath));
            CheckUtils.checkSetting(srcExtension.isPresent() && srcExtension.equals(destExtension),
                "Source (%s) and destination (%s) file formats are not the same",
                srcExtension.map(ExcelFormat::toString).orElse("unknown"),
                destExtension.map(ExcelFormat::toString).orElse("unknown"));
        }

        m_replaceMissings.validateSettings(settings);
        m_missingValPattern.validateSettings(settings);
        m_evaluateFormulas.validateSettings(settings);
    }

    /**
     * Loads the {@link NodeSettingsRO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsRO} to be loaded
     * @throws InvalidSettingsException - If loading the settings failed
     */
    void loadSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_srcFileChooser.loadSettingsFrom(settings);
        m_createNewFile.loadSettingsFrom(settings);
        m_destFileChooser.loadSettingsFrom(settings);
        m_evaluateFormulas.loadSettingsFrom(settings);
        m_replaceMissings.loadSettingsFrom(settings);
        m_missingValPattern.loadSettingsFrom(settings);
        loadSheetNamesForModel(settings);
        loadCoordinateColumnNamesForModel(settings);
    }

    private void loadSheetNamesForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        final String[] sheetNames = settings.getStringArray(CFG_SHEET_NAMES);
        System.arraycopy(sheetNames, 0, m_sheetNames, 0, Math.min(m_sheetNames.length, sheetNames.length));
    }

    private void loadCoordinateColumnNamesForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        final String[] coordinateColumnNames = settings.getStringArray(CFG_COORDINATE_COLUMN_NAMES);
        System.arraycopy(coordinateColumnNames, 0, m_coordinateColumnNames, 0,
            Math.min(m_coordinateColumnNames.length, coordinateColumnNames.length));
    }

    /**
     * Saves the {@link NodeSettingsWO} in the {@link NodeModel}.
     *
     * @param settings the {@link NodeSettingsWO} to save to
     */
    void saveSettingsForModel(final NodeSettingsWO settings) {
        m_srcFileChooser.saveSettingsTo(settings);
        m_createNewFile.saveSettingsTo(settings);
        m_destFileChooser.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        saveSheetNames(settings, m_sheetNames);
        saveCoordinateColumnNames(settings, m_coordinateColumnNames);
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
     * Saves the coordinate column names.
     *
     * @param settings the {@link NodeSettingsWO} to write the coordinate column names to
     * @param array the array to write
     */
    static void saveCoordinateColumnNames(final NodeSettingsWO settings, final String[] array) {
        settings.addStringArray(CFG_COORDINATE_COLUMN_NAMES, array);
    }

    /**
     * Returns the {@link SettingsModelWriterFileChooser}.
     *
     * @return the {@link SettingsModelWriterFileChooser}
     */
    SettingsModelReaderFileChooser getSrcFileChooserModel() {
        return m_srcFileChooser;
    }

    SettingsModelWriterFileChooser getDestFileChooserModel() {
        return m_destFileChooser;
    }

    void setSheetNames(final String[] sheetNames) {
        m_sheetNames = sheetNames;
    }

    void setCoordinateColumnNames(final String[] coordinateColumnNames) {
        m_coordinateColumnNames = coordinateColumnNames;
    }

    /**
     * Loads the sheet names in the {@link NodeDialog}.
     *
     * @param settings the {@link NodeSettingsRO} containing the sheet names
     */
    void loadSheetsInDialog(final NodeSettingsRO settings) {
        try {
            loadSheetNamesForModel(settings);
        } catch (InvalidSettingsException e) { // NOSONAR: just load the default values
            Arrays.fill(m_sheetNames, "");
        }
    }

    /**
     * Loads the coordinate column names in the {@link NodeDialog}.
     *
     * @param settings the {@link NodeSettingsRO} containing the sheet names
     */
    void loadCoordinateColumnNamesInDialog(final NodeSettingsRO settings) {
        try {
            loadCoordinateColumnNamesForModel(settings);
        } catch (InvalidSettingsException e) { // NOSONAR: just load the default values
            Arrays.fill(m_coordinateColumnNames, "");
        }
    }

    @Override
    public String[] getSheetNames() {
        return m_sheetNames;
    }

    String[] getCoordinateColumnNames() {
        return m_coordinateColumnNames;
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
        return false;
    }

    @Override
    public boolean useLandscape() {
        return false;
    }

    @Override
    public short getPaperSize() {
        return 0;
    }

    @Override
    public boolean writeRowKey() {
        return false;
    }

    @Override
    public boolean writeColHeaders() {
        return false;
    }

    @Override
    public boolean abortIfSheetExists() {
        return false;
    }

    @Override
    public boolean evaluate() {
        return m_evaluateFormulas.getBooleanValue();
    }

    @Override
    public ExcelFormat getExcelFormat() {
        return null;
    }

    /**
     * @return the createNewFile
     */
    public SettingsModelBoolean isCreateNewFile() {
        return m_createNewFile;
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

}
