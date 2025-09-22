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
 *   Sep 22, 2025 (AI Migration): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.sheets;

import java.nio.file.Files;
import java.nio.file.Path;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.filehandling.core.data.location.FSLocation;
import org.knime.filehandling.core.data.location.FSLocationSerializationUtils;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.persistence.NodeParametersPersistor;
import org.knime.node.parameters.persistence.Persistor;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileChooserFilters;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.MultiFileSelection;

/**
 * Modern UI settings for the Excel Sheet Reader node. Provides a {@link MultiFileSelection} so the user can either
 * select a single Excel workbook or all Excel workbooks contained in a folder (optionally including subfolders).
 * <p>
 * Backwards compatibility: The persisted settings reproduce the legacy {@code SettingsModelReaderFileChooser}
 * structure under the config key {@code file_selection}, so that the unchanged {@link ExcelSheetReaderNodeModel}
 * continues to work without any modification.
 */
@SuppressWarnings("restriction")
final class ExcelSheetReaderNodeSettings implements NodeParameters {

    /** Supported file extensions (case-insensitive) for filtering purposes. */
    private static final String[] EXCEL_EXTENSIONS = new String[] {".xlsx", ".xlsm", ".xlsb", ".xls"};

    /**
     * Filters for selecting only Excel files when the selection mode is FOLDER. For simplicity we only allow filtering
     * by extension (hidden files are accepted as in the legacy default). Further filter options can be added later
     * without breaking compatibility.
     */
    static final class ExcelFileChooserFilters implements FileChooserFilters {

        @Widget(title = "Follow symlinks", description = "Follow symbolic links while traversing folders (only relevant when selecting a folder).")
        boolean m_followSymlinks = true;

        @Override
        public boolean passesFilter(final Path root, final Path path) {
            if (Files.isDirectory(path)) { // always traverse into directories
                return true;
            }
            final var name = path.getFileName().toString().toLowerCase();
            for (var ext : EXCEL_EXTENSIONS) {
                if (name.endsWith(ext)) {
                    return true;
                }
            }
            return false; // ignore non Excel files
        }

        @Override
        public boolean followSymlinks() {
            return m_followSymlinks;
        }
    }

    /**
     * Multi file selection used in the dialog. It is persisted using {@link FileSelectionPersistor} to be compatible
     * with the legacy settings model key "file_selection".
     */
    @Widget(title = "Excel file selection", description = "Select a single Excel workbook or all Excel workbooks in a folder. Choose 'Type' = File for one file or 'Type' = Folder to include all matching Excel files (optionally from subfolders).")
    @Persistor(FileSelectionPersistor.class)
    MultiFileSelection<ExcelFileChooserFilters> m_fileSelection =
        new MultiFileSelection<>(new ExcelFileChooserFilters());

    /**
     * Persistor that translates the {@link MultiFileSelection} into the legacy settings structure expected by
     * {@code SettingsModelReaderFileChooser} (root key: file_selection).
     */
    public static final class FileSelectionPersistor
        implements NodeParametersPersistor<MultiFileSelection<ExcelFileChooserFilters>> {

        private static final String CFG_KEY = "file_selection";

        @Override
        public MultiFileSelection<ExcelFileChooserFilters> load(final NodeSettingsRO settings)
            throws InvalidSettingsException {
            final var chooserCfg = settings.getNodeSettings(CFG_KEY);

            final var pathCfg = chooserCfg.getNodeSettings("path");
            final FSLocation loc = FSLocationSerializationUtils.loadFSLocation(pathCfg);

            final var filterModeCfg = chooserCfg.getNodeSettings("filter_mode");
            final String mode = filterModeCfg.getString("filter_mode", "FILE");
            final boolean includeSubfolders = filterModeCfg.getBoolean("include_subfolders", false);

            final var selection = new MultiFileSelection<>(new ExcelFileChooserFilters(), loc);
            if ("FILES_IN_FOLDERS".equals(mode)) {
                selection.m_fileOrFolder = MultiFileSelection.FileOrFolder.FOLDER; // NOSONAR
                selection.m_includeSubfolders = includeSubfolders; // NOSONAR
            } else {
                selection.m_fileOrFolder = MultiFileSelection.FileOrFolder.FILE; // NOSONAR
                selection.m_includeSubfolders = false; // NOSONAR
            }
            return selection;
        }

        @Override
        public void save(final MultiFileSelection<ExcelFileChooserFilters> obj, final NodeSettingsWO settings) {
            // Defensive: null object should not happen, but if it does we create an empty selection
            final var selection = obj == null ? new MultiFileSelection<>(new ExcelFileChooserFilters()) : obj;

            final var chooserCfg = settings.addNodeSettings(CFG_KEY);
            // path
            FSLocationSerializationUtils.saveFSLocation(selection.m_path, chooserCfg.addNodeSettings("path"));

            // filter mode & inclusion of subfolders
            final var filterModeCfg = chooserCfg.addNodeSettings("filter_mode");
            final boolean isFolder = selection.m_fileOrFolder == MultiFileSelection.FileOrFolder.FOLDER;
            filterModeCfg.addString("filter_mode", isFolder ? "FILES_IN_FOLDERS" : "FILE");
            filterModeCfg.addBoolean("include_subfolders", isFolder && selection.m_includeSubfolders);

            // filter options (kept minimal – legacy dialog had many options, we default to 'no extra filtering')
            final var filterOptions = filterModeCfg.addNodeSettings("filter_options");
            filterOptions.addBoolean("filter_files_extension", false);
            filterOptions.addString("files_extension_expression", "");
            filterOptions.addBoolean("files_extension_case_sensitive", false);
            filterOptions.addBoolean("filter_files_name", false);
            filterOptions.addString("files_name_expression", "*");
            filterOptions.addBoolean("files_name_case_sensitive", false);
            filterOptions.addString("files_name_filter_type", "WILDCARD");
            filterOptions.addBoolean("include_hidden_files", true); // allow hidden by default
            filterOptions.addBoolean("include_special_files", true);
            filterOptions.addBoolean("filter_folders_name", false);
            filterOptions.addString("folders_name_expression", "*");
            filterOptions.addBoolean("folders_name_case_sensitive", false);
            filterOptions.addString("folders_name_filter_type", "WILDCARD");
            filterOptions.addBoolean("include_hidden_folders", true);
            filterOptions.addBoolean("follow_links", selection.m_filters.m_followSymlinks);

            // internal settings (minimal subset sufficient for the SettingsModelReaderFileChooser)
            final var internals = chooserCfg.addNodeSettings("file_system_chooser__Internals");
            internals.addBoolean("has_fs_port", false); // updated automatically at runtime when ports present
            internals.addBoolean("overwritten_by_variable", false);
            internals.addString("convenience_fs_category", "RELATIVE");
            internals.addString("relative_to", "knime.workflow");
            internals.addString("mountpoint", "LOCAL");
            internals.addString("spaceId", "");
            internals.addString("spaceName", "");
            internals.addInt("custom_url_timeout", 1000);
            internals.addBoolean("connected_fs", true);
        }

        @Override
        public String[][] getConfigPaths() {
            return new String[][] {{CFG_KEY, "path"}};
        }
    }
}
