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
import java.util.Locale;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileChooserFilters;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.MultiFileSelection;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.data.location.FSLocationSerializationUtils;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.persistence.NodeParametersPersistor;
import org.knime.node.parameters.persistence.Persistor;

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

        // ---- File filters -------------------------------------------------

        @Widget(title = "Filter by file extension", description = "Enable filtering files by their extension (e.g. 'xlsx;xlsm'). Use a semicolon-separated list without dots. Case-insensitive unless 'Case sensitive (extensions)' is enabled.")
        boolean m_filterFilesExtension; // legacy: filter_files_extension

        @Widget(title = "Extensions list", description = "Semicolon-separated list of file extensions to include (e.g. 'xlsx;xlsm;xls'). Ignored if 'Filter by file extension' is disabled.")
        String m_filesExtensionExpression = ""; // legacy: files_extension_expression

        @Widget(title = "Case sensitive (extensions)", description = "Treat the entered extensions as case sensitive when matching.")
        boolean m_filesExtensionCaseSensitive; // legacy: files_extension_case_sensitive

        @Widget(title = "Filter by file name", description = "Enable filtering by file name pattern (wildcards '*' and '?' supported) or regular expression depending on filter type.")
        boolean m_filterFilesName; // legacy: filter_files_name

        @Widget(title = "File name pattern", description = "Pattern for file name filtering. With type 'Wildcard', use '*' and '?'. With type 'Regex', enter a Java regular expression.")
        String m_filesNameExpression = "*"; // legacy: files_name_expression

        enum FileNameFilterType { // maps to files_name_filter_type
            /** Wildcard pattern using * and ? */
            WILDCARD,
            /** Java regular expression */
            REGEX;
        }

        @Widget(title = "File name filter type", description = "Choose how to interpret the file name pattern: wildcard or regular expression.")
        FileNameFilterType m_filesNameFilterType = FileNameFilterType.WILDCARD; // legacy: files_name_filter_type

        @Widget(title = "Case sensitive (names)", description = "Make file name filtering case sensitive.")
        boolean m_filesNameCaseSensitive; // legacy: files_name_case_sensitive

        @Widget(title = "Include hidden files", description = "Include hidden files in the selection.")
        boolean m_includeHiddenFiles = true; // legacy: include_hidden_files (we used true previously for convenience)

        @Widget(title = "Include special files", description = "Include special file types (workflows etc).")
        boolean m_includeSpecialFiles = true; // legacy: include_special_files

        // ---- Folder filters -----------------------------------------------

        @Widget(title = "Filter by folder name", description = "Enable filtering of folders by name pattern before descending into them.")
        boolean m_filterFoldersName; // legacy: filter_folders_name

        @Widget(title = "Folder name pattern", description = "Pattern for folder name filtering (same syntax and filter type options as file name pattern). Ignored if 'Filter by folder name' is disabled.")
        String m_foldersNameExpression = "*"; // legacy: folders_name_expression

        enum FolderNameFilterType { // maps to folders_name_filter_type
            WILDCARD,
            REGEX;
        }

        @Widget(title = "Folder name filter type", description = "Choose how to interpret the folder name pattern: wildcard or regular expression.")
        FolderNameFilterType m_foldersNameFilterType = FolderNameFilterType.WILDCARD; // legacy: folders_name_filter_type

        @Widget(title = "Case sensitive (folders)", description = "Make folder name filtering case sensitive.")
        boolean m_foldersNameCaseSensitive; // legacy: folders_name_case_sensitive

        @Widget(title = "Include hidden folders", description = "Descend into folders that are hidden (if they otherwise pass filters).")
        boolean m_includeHiddenFolders = true; // legacy: include_hidden_folders

        // ---- Traversal / symlinks ----------------------------------------

        @Widget(title = "Follow symlinks", description = "Follow symbolic links while traversing folders (only relevant when selecting a folder).")
        boolean m_followSymlinks = true; // legacy: follow_links

        @Override
        public boolean passesFilter(final Path root, final Path path) {
            if (Files.isDirectory(path)) { // always traverse into directories
                return true;
            }
            final var name = path.getFileName().toString().toLowerCase();
            // extension based filtering: if extension filter is enabled, we check user list first; otherwise ensure it's an Excel file
            boolean isExcel = false;
            for (var ext : EXCEL_EXTENSIONS) {
                if (name.endsWith(ext)) { isExcel = true; break; }
            }
            if (!isExcel) {
                return false; // always ignore non-excel files
            }

            if (m_filterFilesExtension && !m_filesExtensionExpression.isBlank()) {
                var allowed = m_filesExtensionExpression.split(";+");
                var fileExt = name.contains(".") ? name.substring(name.lastIndexOf('.') + 1) : "";
                var match = false;
                for (var a : allowed) {
                    var token = a.trim();
                    if (token.isEmpty()) { continue; }
                    if (m_filesExtensionCaseSensitive) {
                        match = fileExt.equals(token);
                    } else {
                        match = fileExt.equalsIgnoreCase(token);
                    }
                    if (match) { break; }
                }
                if (!match) { return false; }
            }

            if (m_filterFilesName) {
                if (!matchesName(name, m_filesNameExpression, m_filesNameFilterType == FileNameFilterType.REGEX,
                        m_filesNameCaseSensitive)) {
                    return false;
                }
            }
            return true;
        }

        @Override
        public boolean followSymlinks() { return m_followSymlinks; }

        private static boolean matchesName(final String filename, final String pattern, final boolean regex,
            final boolean caseSensitive) {
            if (pattern == null || pattern.isEmpty()) { return true; }
            String target = caseSensitive ? filename : filename.toLowerCase(Locale.ROOT);
            String expr = caseSensitive ? pattern : pattern.toLowerCase(Locale.ROOT);
            if (regex) {
                return target.matches(expr);
            }
            // wildcard: convert * -> .* and ? -> . and escape regex specials first
            StringBuilder sb = new StringBuilder();
            for (char c : expr.toCharArray()) {
                switch (c) {
                case '*': sb.append(".*"); break;
                case '?': sb.append('.'); break;
                case '.': case '(': case ')': case '[': case ']': case '{': case '}': case '^': case '$': case '|': case '+': case '\\':
                    sb.append('\\').append(c); break;
                default: sb.append(c);
                }
            }
            return target.matches(sb.toString());
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

            // Attempt to load filter options; if absent (older workflows) keep defaults
            if (filterModeCfg.containsKey("filter_options")) {
                final var filterOptions = filterModeCfg.getNodeSettings("filter_options");
                final var filters = selection.m_filters;
                filters.m_filterFilesExtension = filterOptions.getBoolean("filter_files_extension", false);
                filters.m_filesExtensionExpression = filterOptions.getString("files_extension_expression", "");
                filters.m_filesExtensionCaseSensitive = filterOptions.getBoolean("files_extension_case_sensitive", false);
                filters.m_filterFilesName = filterOptions.getBoolean("filter_files_name", false);
                filters.m_filesNameExpression = filterOptions.getString("files_name_expression", "*");
                filters.m_filesNameCaseSensitive = filterOptions.getBoolean("files_name_case_sensitive", false);
                // file name filter type
                try {
                    filters.m_filesNameFilterType = ExcelFileChooserFilters.FileNameFilterType
                        .valueOf(filterOptions.getString("files_name_filter_type", "WILDCARD"));
                } catch (IllegalArgumentException ex) { // fallback
                    filters.m_filesNameFilterType = ExcelFileChooserFilters.FileNameFilterType.WILDCARD;
                }
                filters.m_includeHiddenFiles = filterOptions.getBoolean("include_hidden_files", true);
                filters.m_includeSpecialFiles = filterOptions.getBoolean("include_special_files", true);
                filters.m_filterFoldersName = filterOptions.getBoolean("filter_folders_name", false);
                filters.m_foldersNameExpression = filterOptions.getString("folders_name_expression", "*");
                filters.m_foldersNameCaseSensitive = filterOptions.getBoolean("folders_name_case_sensitive", false);
                try {
                    filters.m_foldersNameFilterType = ExcelFileChooserFilters.FolderNameFilterType
                        .valueOf(filterOptions.getString("folders_name_filter_type", "WILDCARD"));
                } catch (IllegalArgumentException ex) {
                    filters.m_foldersNameFilterType = ExcelFileChooserFilters.FolderNameFilterType.WILDCARD;
                }
                filters.m_includeHiddenFolders = filterOptions.getBoolean("include_hidden_folders", true);
                filters.m_followSymlinks = filterOptions.getBoolean("follow_links", filters.m_followSymlinks);
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
            final var filters = selection.m_filters;
            filterOptions.addBoolean("filter_files_extension", filters.m_filterFilesExtension);
            filterOptions.addString("files_extension_expression", filters.m_filesExtensionExpression);
            filterOptions.addBoolean("files_extension_case_sensitive", filters.m_filesExtensionCaseSensitive);
            filterOptions.addBoolean("filter_files_name", filters.m_filterFilesName);
            filterOptions.addString("files_name_expression", filters.m_filesNameExpression);
            filterOptions.addBoolean("files_name_case_sensitive", filters.m_filesNameCaseSensitive);
            filterOptions.addString("files_name_filter_type", filters.m_filesNameFilterType.name());
            filterOptions.addBoolean("include_hidden_files", filters.m_includeHiddenFiles);
            filterOptions.addBoolean("include_special_files", filters.m_includeSpecialFiles);
            filterOptions.addBoolean("filter_folders_name", filters.m_filterFoldersName);
            filterOptions.addString("folders_name_expression", filters.m_foldersNameExpression);
            filterOptions.addBoolean("folders_name_case_sensitive", filters.m_foldersNameCaseSensitive);
            filterOptions.addString("folders_name_filter_type", filters.m_foldersNameFilterType.name());
            filterOptions.addBoolean("include_hidden_folders", filters.m_includeHiddenFolders);
            filterOptions.addBoolean("follow_links", filters.m_followSymlinks);

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
