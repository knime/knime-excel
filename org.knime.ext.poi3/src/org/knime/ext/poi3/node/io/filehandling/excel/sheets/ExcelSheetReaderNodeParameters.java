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
import java.nio.file.attribute.BasicFileAttributes;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileChooserFilters;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.MultiFileSelection;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.data.location.FSLocationSerializationUtils;
import org.knime.filehandling.core.defaultnodesettings.filtermode.FileAndFolderFilter;
import org.knime.filehandling.core.defaultnodesettings.filtermode.FilterOptionsSettings;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.layout.Section;
import org.knime.node.parameters.persistence.NodeParametersPersistor;
import org.knime.node.parameters.persistence.Persistor;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;

/**
 * Modern UI settings for the Excel Sheet Reader node. Provides a {@link MultiFileSelection} so the user can either
 * select a single Excel workbook or all Excel workbooks contained in a folder (optionally including subfolders).
 * <p>
 * Backwards compatibility: The persisted settings reproduce the legacy {@code SettingsModelReaderFileChooser} structure
 * under the config key {@code file_selection}, so that the unchanged ExcelSheetReaderNodeModel continues to work
 * without any modification.
 */
@SuppressWarnings("restriction")
final class ExcelSheetReaderNodeParameters implements NodeParameters {

    /**
     * Filters for selecting only Excel files when the selection mode is FOLDER. For simplicity we only allow filtering
     * by extension (hidden files are accepted as in the legacy default). Further filter options can be added later
     * without breaking compatibility.
     */
    static final class DefaultMultiFileChooserFilters implements FileChooserFilters {

        // ---- Sections ----------------------------------------------------
        @Section(title = "File filter options")
        interface FileFilterSection {
        }

        @Section(title = "Folder filter options")
        interface FolderFilterSection {
        }

        @Section(title = "Link options")
        interface LinkOptionsSection {
        }

        // ---- Predicates & References for Effects -------------------------
        // Value reference marker interfaces (must extend ParameterReference<Boolean>)
        interface FileExtensionFilterRef extends ParameterReference<Boolean> {
        }

        interface FileNameFilterRef extends ParameterReference<Boolean> {
        }

        interface FolderNameFilterRef extends ParameterReference<Boolean> {
        }

        // Effect predicate providers
        static final class FileExtensionFilterEnabled implements EffectPredicateProvider {
            @Override
            public EffectPredicate init(final PredicateInitializer i) {
                return i.getBoolean(FileExtensionFilterRef.class).isTrue();
            }
        }

        static final class FileNameFilterEnabled implements EffectPredicateProvider {
            @Override
            public EffectPredicate init(final PredicateInitializer i) {
                return i.getBoolean(FileNameFilterRef.class).isTrue();
            }
        }

        static final class FolderNameFilterEnabled implements EffectPredicateProvider {
            @Override
            public EffectPredicate init(final PredicateInitializer i) {
                return i.getBoolean(FolderNameFilterRef.class).isTrue();
            }
        }

        // ---- File filters -------------------------------------------------

        @Widget(title = "Filter by file extension",
            description = "Enable filtering files by their extension (e.g. 'xlsx;xlsm').")
        @ValueReference(FileExtensionFilterRef.class)
        @Layout(FileFilterSection.class)
        boolean m_filterFilesExtension; // legacy: filter_files_extension

        @Widget(title = "File extensions",
            description = "Semicolon-separated list of file extensions to include (e.g. 'xlsx;xlsm;xls'). Case-insensitive unless 'Case sensitive (extensions)' is enabled.")
        @Layout(FileFilterSection.class)
        @Effect(predicate = FileExtensionFilterEnabled.class, type = EffectType.SHOW)
        String m_filesExtensionExpression = ""; // legacy: files_extension_expression

        @Widget(title = "Case sensitive (extensions)",
            description = "Treat the entered extensions as case sensitive when matching.")
        @Layout(FileFilterSection.class)
        @Effect(predicate = FileExtensionFilterEnabled.class, type = EffectType.SHOW)
        boolean m_filesExtensionCaseSensitive; // legacy: files_extension_case_sensitive

        @Widget(title = "Filter by file name",
            description = "Enable filtering by file name pattern with wildcards or regular expression.")
        @ValueReference(FileNameFilterRef.class)
        @Layout(FileFilterSection.class)
        boolean m_filterFilesName; // legacy: filter_files_name

        @Widget(title = "File name filter pattern",
            description = "Pattern for file name filtering. With type 'Wildcard', use '*' and '?'. With type 'Regex', enter a Java regular expression.")
        @Layout(FileFilterSection.class)
        @Effect(predicate = FileNameFilterEnabled.class, type = EffectType.SHOW)
        String m_filesNameExpression = "*"; // legacy: files_name_expression

        enum NameFilterType {
                WILDCARD, REGEX;
        }

        @Widget(title = "File name filter type",
            description = "Choose how to interpret the file name pattern: wildcard or regular expression.")
        @ValueSwitchWidget
        @Layout(FileFilterSection.class)
        @Effect(predicate = FileNameFilterEnabled.class, type = EffectType.SHOW)
        NameFilterType m_filesNameFilterType = NameFilterType.WILDCARD; // legacy: files_name_filter_type

        @Widget(title = "Case sensitive (names)", description = "Make file name filtering case sensitive.")
        @Layout(FileFilterSection.class)
        @Effect(predicate = FileNameFilterEnabled.class, type = EffectType.SHOW)
        boolean m_filesNameCaseSensitive; // legacy: files_name_case_sensitive

        @Widget(title = "Include hidden files", description = "Include hidden files in the selection.")
        @Layout(FileFilterSection.class)
        boolean m_includeHiddenFiles = true; // legacy: include_hidden_files

        @Widget(title = "Include special files", description = "Include special file types (workflows etc).")
        @Layout(FileFilterSection.class)
        boolean m_includeSpecialFiles = true; // legacy: include_special_files

        // ---- Folder filters -----------------------------------------------

        @Widget(title = "Filter by folder name",
            description = "Enable filtering of folders by name pattern before descending into them.")
        @ValueReference(FolderNameFilterRef.class)
        @Layout(FolderFilterSection.class)
        boolean m_filterFoldersName; // legacy: filter_folders_name

        @Widget(title = "Folder name pattern", description = """
                Pattern for folder name filtering. Use '*' and '?' with filter type 'Wildcard'.
                With type 'Regex', enter a Java regular expression.
                """)
        @Layout(FolderFilterSection.class)
        @Effect(predicate = FolderNameFilterEnabled.class, type = EffectType.ENABLE)
        String m_foldersNameExpression = "*"; // legacy: folders_name_expression

        @Widget(title = "Folder name filter type",
            description = "Choose how to interpret the folder name pattern: wildcard or regular expression.")
        @ValueSwitchWidget
        @Layout(FolderFilterSection.class)
        @Effect(predicate = FolderNameFilterEnabled.class, type = EffectType.ENABLE)
        NameFilterType m_foldersNameFilterType = NameFilterType.WILDCARD; // legacy: folders_name_filter_type

        @Widget(title = "Case sensitive (folders)", description = "Make folder name filtering case sensitive.")
        @Layout(FolderFilterSection.class)
        @Effect(predicate = FolderNameFilterEnabled.class, type = EffectType.ENABLE)
        boolean m_foldersNameCaseSensitive; // legacy: folders_name_case_sensitive

        @Widget(title = "Include hidden folders",
            description = "Descend into folders that are hidden (if they otherwise pass filters).")
        @Layout(FolderFilterSection.class)
        boolean m_includeHiddenFolders = true; // legacy: include_hidden_folders

        @Widget(title = "Follow symlinks",
            description = "Follow symbolic links while traversing folders (only relevant when selecting a folder).")
        @Layout(LinkOptionsSection.class)
        boolean m_followSymlinks = true; // legacy: follow_links

        @Override
        public boolean passesFilter(final Path root, final Path path) {
            // We're constructing a FileAndFolderFilter here, as that is what is also used
            // in the model and was used in the old dialog. This way we ensure the same
            // filtering behavior in the dialog preview as in the model.
            FileAndFolderFilter filter = new FileAndFolderFilter(root, toFilterOptionsSettings());
            try {
                BasicFileAttributes attrs = Files.readAttributes(path, BasicFileAttributes.class);
                return filter.test(path, attrs);
            } catch (Exception e) {
                // If we can't read attributes, fallback to legacy logic: skip file
                return false;
            }
        }

        public FilterOptionsSettings toFilterOptionsSettings() {
            FilterOptionsSettings filterOptions = new FilterOptionsSettings();

            filterOptions.setFilterFilesByName(m_filterFilesName);
            filterOptions.setFilesNameCaseSensitive(m_filesNameCaseSensitive);
            filterOptions.setFilesNameExpression(m_filesNameExpression);
            filterOptions.setFilesNameFilterType(m_filesNameFilterType.equals(NameFilterType.WILDCARD)
                ? FileAndFolderFilter.FilterType.WILDCARD : FileAndFolderFilter.FilterType.REGEX);

            filterOptions.setFilterFilesByExtension(m_filterFilesExtension);
            filterOptions.setFilesExtensionCaseSensitive(m_filesExtensionCaseSensitive);
            filterOptions.setFilesExtensionExpression(m_filesExtensionExpression);

            filterOptions.setFilterFoldersByName(m_filterFoldersName);
            filterOptions.setFoldersNameCaseSensitive(m_foldersNameCaseSensitive);
            filterOptions.setFoldersNameExpression(m_foldersNameExpression);
            filterOptions.setFoldersNameFilterMode(m_foldersNameFilterType.equals(NameFilterType.WILDCARD)
                ? FileAndFolderFilter.FilterType.WILDCARD : FileAndFolderFilter.FilterType.REGEX);

            filterOptions.setIncludeSpecialFiles(m_includeSpecialFiles);
            filterOptions.setFollowLinks(m_followSymlinks);
            filterOptions.setIncludeHiddenFolders(m_includeHiddenFolders);
            filterOptions.setIncludeHiddenFiles(m_includeHiddenFiles);
            return filterOptions;
        }

        public void setFromFilterOptionsSettings(final FilterOptionsSettings filterOptions) {
            m_filterFilesExtension = filterOptions.isFilterFilesByExtension();
            m_filesExtensionExpression = filterOptions.getFilesExtensionExpression();
            m_filesExtensionCaseSensitive = filterOptions.isFilesExtensionCaseSensitive();
            m_filterFilesName = filterOptions.isFilterFilesByName();
            m_filesNameExpression = filterOptions.getFilesNameExpression();
            m_filesNameCaseSensitive = filterOptions.isFilesNameCaseSensitive();
            // file name filter type
            try {
                m_filesNameFilterType = DefaultMultiFileChooserFilters.NameFilterType
                    .valueOf(filterOptions.getFilesNameFilterType().name());
            } catch (IllegalArgumentException ex) { // fallback
                m_filesNameFilterType = DefaultMultiFileChooserFilters.NameFilterType.WILDCARD;
            }
            m_includeHiddenFiles = filterOptions.isIncludeHiddenFiles();
            m_includeSpecialFiles = filterOptions.isIncludeSpecialFiles();
            m_filterFoldersName = filterOptions.isFilterFoldersByName();
            m_foldersNameExpression = filterOptions.getFoldersNameExpression();
            m_foldersNameCaseSensitive = filterOptions.isFoldersNameCaseSensitive();
            try {
                m_foldersNameFilterType = DefaultMultiFileChooserFilters.NameFilterType
                    .valueOf(filterOptions.getFoldersNameFilterType().name());
            } catch (IllegalArgumentException ex) {
                m_foldersNameFilterType = DefaultMultiFileChooserFilters.NameFilterType.WILDCARD;
            }
            m_includeHiddenFolders = filterOptions.isIncludeHiddenFolders();
            m_followSymlinks = filterOptions.isFollowLinks();
        }

        @Override
        public boolean followSymlinks() {
            return m_followSymlinks;
        }

    }

    @Widget(title = "Excel file selection", description = """
            Select a single Excel workbook or all Excel workbooks in a folder. Choose 'Type' = File for
            one file or 'Type' = Folder to include all matching Excel files (optionally from subfolders).
            """)
    @Persistor(ExcelFileSelectionPersistor.class)
    MultiFileSelection<DefaultMultiFileChooserFilters> m_fileSelection =
        new MultiFileSelection<>(new DefaultMultiFileChooserFilters());

    /**
     * Persistor that translates the {@link MultiFileSelection} into the legacy settings structure expected by
     * {@code SettingsModelReaderFileChooser} (root key: file_selection).
     */
    static abstract class LegacyMultiFileReaderPersistor
        implements NodeParametersPersistor<MultiFileSelection<DefaultMultiFileChooserFilters>> {

        private final String m_configKey;

        LegacyMultiFileReaderPersistor(final String configKey) {
            m_configKey = configKey;
        }

        @Override
        public MultiFileSelection<DefaultMultiFileChooserFilters> load(final NodeSettingsRO settings)
            throws InvalidSettingsException {
            final var chooserCfg = settings.getNodeSettings(m_configKey);

            final var pathCfg = chooserCfg.getNodeSettings("path");
            final FSLocation loc = FSLocationSerializationUtils.loadFSLocation(pathCfg);

            final var filterModeCfg = chooserCfg.getNodeSettings("filter_mode");
            final String mode = filterModeCfg.getString("filter_mode", "FILE");
            final boolean includeSubfolders = filterModeCfg.getBoolean("include_subfolders", false);

            final var fileSelection = new MultiFileSelection<>(new DefaultMultiFileChooserFilters(), loc);
            if ("FILES_IN_FOLDERS".equals(mode)) {
                fileSelection.m_fileOrFolder = MultiFileSelection.FileOrFolder.FOLDER; // NOSONAR
                fileSelection.m_includeSubfolders = includeSubfolders; // NOSONAR
            } else {
                fileSelection.m_fileOrFolder = MultiFileSelection.FileOrFolder.FILE; // NOSONAR
                fileSelection.m_includeSubfolders = false; // NOSONAR
            }

            // Attempt to load filter options; if absent (older workflows) keep defaults
            if (filterModeCfg.containsKey("filter_options")) {
                final var filterOptions = filterModeCfg.getNodeSettings("filter_options");
                final var filterOptionsSettings = new FilterOptionsSettings();
                filterOptionsSettings.loadFromConfigForDialog(filterOptions);
                fileSelection.m_filters.setFromFilterOptionsSettings(filterOptionsSettings);
            }
            return fileSelection;
        }

        @Override
        public void save(final MultiFileSelection<DefaultMultiFileChooserFilters> obj, final NodeSettingsWO settings) {
            // Defensive: null object should not happen, but if it does we create an empty selection
            final var fileSelection =
                obj == null ? new MultiFileSelection<>(new DefaultMultiFileChooserFilters()) : obj;

            final var chooserCfg = settings.addNodeSettings(m_configKey);

            // path
            FSLocationSerializationUtils.saveFSLocation(fileSelection.m_path, chooserCfg.addNodeSettings("path"));

            // filter mode & inclusion of subfolders
            final var filterModeCfg = chooserCfg.addNodeSettings("filter_mode");
            final boolean isFolder = fileSelection.m_fileOrFolder == MultiFileSelection.FileOrFolder.FOLDER;
            filterModeCfg.addString("filter_mode", isFolder ? "FILES_IN_FOLDERS" : "FILE");
            filterModeCfg.addBoolean("include_subfolders", isFolder && fileSelection.m_includeSubfolders);

            final var filterOptions = filterModeCfg.addNodeSettings("filter_options");

            final var filterOptionsSettings = fileSelection.m_filters.toFilterOptionsSettings();
            filterOptionsSettings.saveToConfig(filterOptions);

            // internal settings
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
            return new String[][]{//
                {m_configKey, "path"}, //
                {m_configKey, "filter_mode", "filter_mode"}, //
                {m_configKey, "filter_mode", "include_subfolders"}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_FILTER_BY_EXTENSION}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_EXTENSION_EXPRESSION}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_EXTENSION_CASE_SENSITIVE}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_FILTER_BY_NAME}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_NAME_EXPRESSION}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_NAME_CASE_SENSITIVE}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FILES_NAME_FILTER_TYPE}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_INCLUDE_HIDDEN_FILES}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_INCLUDE_SPECIAL_FILES}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FOLDERS_FILTER_BY_NAME}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FOLDERS_NAME_EXPRESSION}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FOLDERS_NAME_CASE_SENSITIVE}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FOLDERS_NAME_FILTER_TYPE}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_INCLUDE_HIDDEN_FOLDERS}, //
                {m_configKey, "filter_options", FilterOptionsSettings.CFG_FOLLOW_LINKS}, //
                {m_configKey, "file_system_chooser__Internals", "has_fs_port"}, //
                {m_configKey, "file_system_chooser__Internals", "overwritten_by_variable"}, //
                {m_configKey, "file_system_chooser__Internals", "convenience_fs_category"}, //
                {m_configKey, "file_system_chooser__Internals", "relative_to"}, //
                {m_configKey, "file_system_chooser__Internals", "mountpoint"}, //
                {m_configKey, "file_system_chooser__Internals", "spaceId"}, //
                {m_configKey, "file_system_chooser__Internals", "spaceName"}, //
                {m_configKey, "file_system_chooser__Internals", "custom_url_timeout"}, //
                {m_configKey, "file_system_chooser__Internals", "connected_fs"},//
            };
        }
    }

    static final class ExcelFileSelectionPersistor extends LegacyMultiFileReaderPersistor {
        ExcelFileSelectionPersistor() {
            super("file_selection");
        }
    }
}
