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
 * ------------------------------------------------------------------------
 */

package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import java.nio.file.Files;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.function.Supplier;
import java.util.stream.IntStream;

import org.apache.poi.EncryptedDocumentException;
import org.knime.base.node.io.filehandling.webui.FileChooserPathAccessor;
import org.knime.base.node.io.filehandling.webui.FileSystemPortConnectionUtil;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileSelection;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileWriterWidget;
import org.knime.core.webui.node.dialog.defaultdialog.persistence.booleanhelpers.DoNotPersistBoolean;
import org.knime.core.webui.node.dialog.defaultdialog.util.updates.StateComputationFailureException;
import org.knime.core.webui.node.dialog.defaultdialog.widget.Modification;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelEncryptionSettings;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.array.ArrayWidget;
import org.knime.node.parameters.array.ArrayWidget.ElementLayout;
import org.knime.node.parameters.array.PerPortValueProvider;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.layout.Section;
import org.knime.node.parameters.migration.LoadDefaultsForAbsentFields;
import org.knime.node.parameters.migration.Migrate;
import org.knime.node.parameters.persistence.NodeParametersPersistor;
import org.knime.node.parameters.persistence.Persist;
import org.knime.node.parameters.persistence.Persistor;
import org.knime.node.parameters.persistence.legacy.LegacyFileWriterWithOverwritePolicyOptions;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.StateProvider;
import org.knime.node.parameters.updates.ValueProvider;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.updates.legacy.LegacyPredicateInitializer;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.StringChoicesProvider;
import org.knime.node.parameters.widget.choices.SuggestionsProvider;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.message.TextMessage;
import org.knime.node.parameters.widget.message.TextMessage.MessageType;
import org.knime.node.parameters.widget.message.TextMessage.SimpleTextMessageProvider;
import org.knime.node.parameters.widget.text.TextInputWidget;

/**
 * Node parameters for Excel Writer.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline v1.2
 */
@SuppressWarnings("restriction")
@LoadDefaultsForAbsentFields
class ExcelTableWriterNodeParameters implements NodeParameters {

    @Section(title = "Output")
    interface OutputSection {
    }

    @Section(title = "Sheets")
    @After(OutputSection.class)
    interface SheetsSection {
    }

    @Section(title = "Headers / Keys")
    @Advanced
    @After(SheetsSection.class)
    interface NamesAndIdsSection {
    }

    @Section(title = "Values")
    @Advanced
    @After(NamesAndIdsSection.class)
    interface ValuesSection {
    }

    @Section(title = "Layout")
    @Advanced
    @After(ValuesSection.class)
    interface LayoutSection {
    }

    @Section(title = "Interaction")
    @Advanced
    @After(LayoutSection.class)
    interface InteractionSection {
    }

    @Widget(title = "Excel format", description = "Select the Excel file format to write.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_EXCEL_FORMAT)
    @Layout(OutputSection.class)
    @ValueSwitchWidget
    @ValueReference(ExcelFormatRef.class)
    ExcelFormat m_excelFormat = ExcelFormat.XLSX;

    @Modification(OutputFileModification.class)
    @Layout(OutputSection.class)
    @Persist(configKey = ExcelTableWriterConfig.CFG_FILE_CHOOSER)
    @ValueReference(OutputFileRef.class)
    LegacyFileWriterWithOverwritePolicyOptions m_outputFile = new LegacyFileWriterWithOverwritePolicyOptions();

    @Widget(title = "Password to protect files",
        description = "If enabled, the output file will be password-protected. "
            + "This supports XLS files (using weak encryption) and XLSX files (using AES encryption).")
    @Layout(OutputSection.class)
    @Persist(configKey = ExcelTableWriterConfig.CFG_AUTHENTICATION_METHOD)
    @ValueReference(EncryptionRef.class)
    ExcelEncryptionSettings m_encryption = new ExcelEncryptionSettings();

    @TextMessage(EncryptionValidationMessageProvider.class)
    @Layout(OutputSection.class)
    @Effect(predicate = OutputFileOverwritePolicyIsAppend.class, type = Effect.EffectType.SHOW)
    Void m_encryptionValidationMessage;

    @Widget(title = "Sheets", description = "")
    @ArrayWidget(hasFixedSize = true, elementDefaultValueProvider = SheetNamesDefaultProvider.class,
        elementLayout = ElementLayout.HORIZONTAL_SINGLE_LINE)
    @ValueProvider(SheetNamesValueProvider.class)
    @Persistor(SheetNamesPersistor.class)
    @Migrate
    @Layout(SheetsSection.class)
    @ValueReference(SheetNamesParameterReference.class)
    SheetName[] m_sheetNames = new SheetName[0];

    @Widget(title = "If sheet exists",
        description = "Specify the behavior of the node in case a sheet with the entered name already exists. "
            + "(This option is only relevant if the file append option is selected.)")
    @ValueSwitchWidget
    @Persist(configKey = ExcelTableWriterConfig.CFG_SHEET_EXISTS)
    @Layout(SheetsSection.class)
    @Effect(predicate = OutputFileOverwritePolicyIsAppend.class, type = Effect.EffectType.SHOW)
    @ValueReference(SheetExistsPolicyRef.class)
    @ValueProvider(SheetExistsPolicyFromConcatenateSheetsUpdater.class)
    SheetExistsPolicy m_ifSheetExists = SheetExistsPolicy.FAIL;

    @Widget(title = "Merge data with identical sheet names into one sheet",
        description = "If checked, it is possible to provide the same sheet name for multiple input tables. "
            + "Data from those tables will be concatenated into the same sheet in the order the tables are connected.")
    @Layout(SheetsSection.class)
    @Effect(predicate = OutputFileOverwritePolicyIsAppend.class, type = Effect.EffectType.HIDE)
    @ValueReference(ConcatenateSheetsWithSameNameRef.class)
    @ValueProvider(ConcatenateSheetsWithSameNameUpdater.class)
    @Persistor(DoNotPersistBoolean.class)
    boolean m_concatenateSheetsWithSameName;

    @TextMessage(SheetNameValidationMessageProvider.class)
    @Layout(SheetsSection.class)
    Void m_sheetNameValidationMessage;

    @Widget(title = "Write row key",
        description = "If checked, the row IDs are added to the output, in the first column of the spreadsheet.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_WRITE_ROW_KEY)
    @Layout(NamesAndIdsSection.class)
    boolean m_writeRowKey;

    @Widget(title = "Write column headers",
        description = "If checked, the column names are written out in the first row of the spreadsheet.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_WRITE_COLUMN_HEADER)
    @Layout(NamesAndIdsSection.class)
    boolean m_writeColumnHeaders = true;

    @Widget(title = "Don't write column headers if sheet exists",
        description = "Only write the column headers if a sheet is newly created or replaced. "
            + "This option is convenient if you have written data with the same specification to an existing sheet "
            + "before and want to append new rows to it.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_SKIP_COLUMN_HEADER_ON_APPEND)
    @Layout(NamesAndIdsSection.class)
    boolean m_skipColumnHeaderOnAppend = true;

    @Widget(title = "Replace missing values",
        description = "If selected, missing values will be replaced by the specified value, "
            + "otherwise a blank cell is being created.")
    @Persistor(MissingValuePatternPersistor.class)
    @Migrate
    @Layout(ValuesSection.class)
    @TextInputWidget(placeholder = "Replacement value")
    Optional<String> m_missingValuePattern = Optional.empty();

    @Widget(title = "Evaluate formulas (leave unchecked if uncertain, see help for details)",
        description = "If checked, all formulas in the file will be evaluated after the sheet has been written. "
            + "This is useful if other sheets in the file refer to the data just written and their content needs "
            + "updating. This option is only relevant if the append option is selected. This can cause errors when "
            + "there are functions that are not implemented by the underlying Apache POI library. "
            + "Note: For xlsx files, evaluation requires significantly more memory as the whole file needs to be kept "
            + "in memory (xls files are anyway loaded completely into memory).")
    @Persist(configKey = ExcelTableWriterConfig.CFG_EVALUATE_FORMULAS)
    @Layout(ValuesSection.class)
    boolean m_evaluateFormulas;

    @Widget(title = "Autosize columns", description = "Fits each column's width to its content.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_AUTOSIZE)
    @Layout(LayoutSection.class)
    boolean m_autosizeColumns;

    @Widget(title = "Page orientation", description = "Sets the print format to portrait or landscape.")
    @ValueSwitchWidget
    @Persist(configKey = ExcelTableWriterConfig.CFG_LANDSCAPE)
    @Layout(LayoutSection.class)
    Orientation m_orientation = Orientation.PORTRAIT;

    @Widget(title = "Paper size", description = "Sets the paper size in the print setup.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_PAPER_SIZE)
    @Layout(LayoutSection.class)
    PaperSize m_paperSize = PaperSize.A4_PAPERSIZE;

    @Layout(InteractionSection.class)
    @Widget(title = "Open file after execution",
        description = "If enabled, the output file will be opened in the associated application "
            + "after the node has successfully executed.")
    @Persist(configKey = ExcelTableWriterConfig.CFG_OPEN_FILE_AFTER_EXEC)
    boolean m_openOutputFileAfterExecution;

    private static final class OutputFileModification implements LegacyFileWriterWithOverwritePolicyOptions.Modifier {
        private static final class ExcelOverwritePolicyChoicesProvider
            extends LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicyChoicesProvider {

            @Override
            protected List<LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy> getChoices() {
                return List.of(LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.fail,
                    LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.overwrite,
                    LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.append);
            }
        }

        @Override
        public void modify(final Modification.WidgetGroupModifier group) {
            restrictOverwritePolicyOptions(group, ExcelOverwritePolicyChoicesProvider.class);
            var fileSelection = findFileSelection(group);
            fileSelection //
                .addAnnotation(ValueReference.class) //
                .withValue(OutputFileSelectionRef.class) //
                .modify();
            fileSelection //
                .addAnnotation(ValueProvider.class) //
                .withValue(OutputLocationExtensionChanger.class) //
                .modify();
            fileSelection //
                .modifyAnnotation(FileWriterWidget.class) //
                .withProperty("fileExtensionProvider", OutputLocationFileExtensionProvider.class) //
                .modify();
            findOverwritePolicy(group) //
                .addAnnotation(ValueReference.class) //
                .withValue(OutputFileOverwritePolicyRef.class) //
                .modify();
        }
    }

    private static class OutputFileSelectionRef implements ParameterReference<FileSelection> {
    }

    private static class OutputFileOverwritePolicyRef
        implements ParameterReference<LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy> {
    }

    private static final class OutputFileRef implements ParameterReference<LegacyFileWriterWithOverwritePolicyOptions> {
    }

    private static final class OutputFileOverwritePolicyIsAppend implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return ((LegacyPredicateInitializer)i).getLegacyFileWriter(OutputFileRef.class).getOverwritePolicy()
                .isOneOf(LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.append);
        }
    }

    private static final class OutputLocationExtensionChanger implements StateProvider<FileSelection> {

        Supplier<ExcelFormat> m_format;

        Supplier<FileSelection> m_outputLocation;

        @Override
        public void init(final StateProviderInitializer initializer) {
            m_format = initializer.computeFromValueSupplier(ExcelFormatRef.class);
            m_outputLocation = initializer.getValueSupplier(OutputFileSelectionRef.class);
        }

        @Override
        public FileSelection computeState(final NodeParametersInput parametersInput)
            throws StateComputationFailureException {
            final var oldOutputLocation = m_outputLocation.get();
            final var oldFsLocation = oldOutputLocation.getFSLocation();
            final var oldPath = oldFsLocation.getPath();
            final var format = m_format.get();
            final var newExtension = "." + format.name().toLowerCase(Locale.ROOT);

            // if path ends with any of the known extensions, replace it
            for (final var knownFormat : ExcelFormat.values()) {
                final var knownExtension = "." + knownFormat.name().toLowerCase(Locale.ROOT);
                if (oldPath.endsWith(knownExtension)) {
                    final var newPath = oldPath.substring(0, oldPath.length() - knownExtension.length()) + newExtension;
                    final var newLocation = new FSLocation(oldFsLocation.getFSCategory(),
                        oldFsLocation.getFileSystemSpecifier().orElse(null), newPath);
                    return new FileSelection(newLocation);
                }
            }
            throw new StateComputationFailureException();
        }
    }

    private static final class OutputLocationFileExtensionProvider implements StateProvider<String> {

        Supplier<ExcelFormat> m_format;

        @Override
        public void init(final StateProviderInitializer initializer) {
            initializer.computeBeforeOpenDialog();
            m_format = initializer.computeFromValueSupplier(ExcelFormatRef.class);
        }

        @Override
        public String computeState(final NodeParametersInput parametersInput) throws StateComputationFailureException {
            final var format = m_format.get();
            return format.name().toLowerCase(Locale.ROOT);
        }
    }

    private static final class SheetNamesParameterReference implements ParameterReference<SheetName[]> {
    }

    private static class SheetName implements NodeParameters {

        SheetName() {
            // Default constructor
        }

        SheetName(final String sheetName) {
            m_sheetName = sheetName;
        }

        @Widget(title = "Sheet name",
            description = "Name of the spreadsheets that will be created. The dropdown can be used to select a sheet "
                + "name which already exists in the Excel file or a custom name can be entered. If \"If sheet exists\" "
                + "isn't set to append, each sheet name must be unique. The node appends the tables in the order they "
                + "are connected.")
        @SuggestionsProvider(SheetNamesChoicesProvider.class)
        String m_sheetName = "Sheet1";
    }

    static class SheetNamesDefaultProvider implements StateProvider<String[]> {

        @Override
        public void init(final StateProviderInitializer initializer) {
            // Nothing to initialize
        }

        @Override
        public String[] computeState(final NodeParametersInput context) {
            final int numSheets = context.getPortsConfiguration().getInputPortLocation()
                .get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length;
            return IntStream.range(0, numSheets).mapToObj(ExcelTableWriterConfig::createDefaultSheetName)
                .toArray(String[]::new);
        }
    }

    private static class SheetNamesPersistor implements NodeParametersPersistor<SheetName[]> {

        @Override
        public SheetName[] load(final NodeSettingsRO settings) throws InvalidSettingsException {
            var sheetNames = settings.getStringArray(ExcelTableWriterConfig.CFG_SHEET_NAMES);
            var result = new SheetName[sheetNames.length];
            for (var i = 0; i < sheetNames.length; i++) {
                result[i] = new SheetName(sheetNames[i]);
            }

            return result;
        }

        @Override
        public void save(final SheetName[] param, final NodeSettingsWO settings) {
            var sheetNames = new String[param.length];
            for (var i = 0; i < param.length; i++) {
                sheetNames[i] = param[i].m_sheetName;
            }
            settings.addStringArray(ExcelTableWriterConfig.CFG_SHEET_NAMES, sheetNames);
        }

        @Override
        public String[][] getConfigPaths() {
            return new String[][]{{ExcelTableWriterConfig.CFG_SHEET_NAMES}};
        }

    }

    private static class SheetNamesValueProvider extends PerPortValueProvider<SheetName> {

        SheetNamesValueProvider() {
            super(ExcelTableWriterNodeFactory.SHEET_GRP_ID, PortGroupSide.INPUT);
        }

        @Override
        protected Supplier<SheetName[]> supplier(final StateProviderInitializer initializer) {
            return initializer.getValueSupplier(SheetNamesParameterReference.class);
        }

        @Override
        protected SheetName[] newArray(final int size) {
            return new SheetName[size];
        }

        @Override
        protected SheetName newInstance(final int index) {
            return new SheetName("default_" + (index + 1));
        }
    }

    private static class SheetNamesChoicesProvider implements StringChoicesProvider {

        private static final NodeLogger LOGGER = NodeLogger.getLogger(SheetNamesChoicesProvider.class);

        Supplier<FileSelection> m_outputFileSelection;

        Supplier<ExcelEncryptionSettings> m_encryption;

        @Override
        public void init(final StateProviderInitializer initializer) {
            /*
             * Looking up sheet names can take time, especially for large files or remote file systems
             * -> do it after opening the dialog (otherwise the dialog opening is blocked)
             */
            initializer.computeAfterOpenDialog();
            m_outputFileSelection = initializer.computeFromValueSupplier(OutputFileSelectionRef.class);
            m_encryption = initializer.computeFromValueSupplier(EncryptionRef.class);
        }

        @Override
        public List<String> choices(final NodeParametersInput context) {
            var fsConnection = FileSystemPortConnectionUtil.getFileSystemConnection(context);
            try (final var accessor = new FileChooserPathAccessor(m_outputFileSelection.get(), fsConnection)) {
                final var path = accessor.getPaths(s -> {
                }).get(0);
                if (!Files.exists(path)) {
                    LOGGER.info("File %s does not exist, cannot provide sheet name suggestions.".formatted(path));
                    return Collections.emptyList();
                }
                try (final var inputStream = FSFiles.newInputStream(path)) {
                    final var password = m_encryption.get().getPassword(context);
                    return ExcelUtils.readSheetNames(inputStream, password);
                }
            } catch (final Exception e) {
                // in case of any error, return no choices
                LOGGER.warn("Could not read sheet names for suggestions.", e);
                return Collections.emptyList();
            }
        }
    }

    private static class EncryptionValidationMessageProvider implements SimpleTextMessageProvider {

        Supplier<FileSelection> m_outputFileSelection;

        Supplier<ExcelEncryptionSettings> m_encryption;

        @Override
        public void init(final StateProviderInitializer initializer) {
            initializer.computeBeforeOpenDialog();
            m_outputFileSelection = initializer.computeFromValueSupplier(OutputFileSelectionRef.class);
            m_encryption = initializer.computeFromValueSupplier(EncryptionRef.class);
        }

        @Override
        public boolean showMessage(final NodeParametersInput context) {
            var fsConnection = FileSystemPortConnectionUtil.getFileSystemConnection(context);
            try (final var accessor = new FileChooserPathAccessor(m_outputFileSelection.get(), fsConnection)) {
                final var path = accessor.getPaths(s -> {
                }).get(0);
                if (!Files.exists(path)) {
                    return false;
                }
                try (final var inputStream = FSFiles.newInputStream(path)) {
                    final var password = m_encryption.get().getPassword(context);
                    CryptUtil.verifyPassword(inputStream, password);
                    return false;
                } catch (final EncryptedDocumentException e) {
                    return true;
                }
            } catch (final Exception e) {
                // in case of any error, do not show the message
                return false;
            }
        }

        @Override
        public String title() {
            return "Invalid password";
        }

        @Override
        public String description() {
            return "The provided password is not valid for the selected file. Please provide a correct password"
                + " via the \"Encryption\" settings.";
        }

        @Override
        public MessageType type() {
            return MessageType.INFO;
        }
    }

    private static class SheetNameValidationMessageProvider implements SimpleTextMessageProvider {

        Supplier<LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy> m_overwritePolicy;

        Supplier<SheetName[]> m_sheetNames;

        Supplier<SheetExistsPolicy> m_ifSheetExists;

        @Override
        public void init(final StateProviderInitializer initializer) {
            initializer.computeBeforeOpenDialog();
            m_overwritePolicy = initializer.computeFromValueSupplier(OutputFileOverwritePolicyRef.class);
            m_sheetNames = initializer.computeFromValueSupplier(SheetNamesParameterReference.class);
            m_ifSheetExists = initializer.computeFromValueSupplier(SheetExistsPolicyRef.class);
        }

        @Override
        public boolean showMessage(final NodeParametersInput context) {
            if (m_ifSheetExists.get() == SheetExistsPolicy.APPEND) {
                return false;
            }
            var sheetNames = m_sheetNames.get();
            var nameSet = Collections.newSetFromMap(new java.util.HashMap<String, Boolean>());
            for (var sheetName : sheetNames) {
                if (!nameSet.add(sheetName.m_sheetName)) {
                    return true;
                }
            }
            return false;
        }

        @Override
        public String title() {
            return "Duplicate sheet names";
        }

        @Override
        public String description() {
            if (m_overwritePolicy.get() == LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.append) {
                return "Please rename the sheets to have unique names or set the \"If sheet exists\" option to "
                    + "\"Append\" to allow duplicate sheet names.";
            } else {
                return "Please rename the sheets to have unique names or enable the \"Merge data with identical sheet "
                    + "names into one sheet\" option to allow duplicate sheet names.";
            }
        }

        @Override
        public MessageType type() {
            return MessageType.INFO;
        }
    }

    private static class ConcatenateSheetsWithSameNameRef implements ParameterReference<Boolean> {
    }

    private static class SheetExistsPolicyFromConcatenateSheetsUpdater implements StateProvider<SheetExistsPolicy> {

        Supplier<Boolean> m_concatenateSheetsWithSameName;

        Supplier<SheetExistsPolicy> m_ifSheetExists;

        @Override
        public void init(final StateProviderInitializer initializer) {
            m_concatenateSheetsWithSameName =
                initializer.computeFromValueSupplier(ConcatenateSheetsWithSameNameRef.class);
            m_ifSheetExists = initializer.getValueSupplier(SheetExistsPolicyRef.class);
        }

        @Override
        public SheetExistsPolicy computeState(final NodeParametersInput parametersInput)
            throws StateComputationFailureException {
            if (Boolean.TRUE.equals(m_concatenateSheetsWithSameName.get())) {
                return SheetExistsPolicy.APPEND;
            }
            if (m_ifSheetExists.get() == SheetExistsPolicy.APPEND) {
                return SheetExistsPolicy.FAIL;
            }
            // Do not overwrite the existing value if it is compatible with the current state
            throw new StateComputationFailureException();
        }
    }

    private static class ConcatenateSheetsWithSameNameUpdater implements StateProvider<Boolean> {
        Supplier<SheetExistsPolicy> m_ifSheetExists;

        @Override
        public void init(final StateProviderInitializer initializer) {
            initializer.computeBeforeOpenDialog();
            initializer.computeOnValueChange(OutputFileRef.class);
            m_ifSheetExists = initializer.getValueSupplier(SheetExistsPolicyRef.class);
        }

        @Override
        public Boolean computeState(final NodeParametersInput parametersInput) {
            return m_ifSheetExists.get() == SheetExistsPolicy.APPEND;
        }
    }

    private static class ExcelFormatRef implements ParameterReference<ExcelFormat> {
    }

    enum ExcelFormat {
            @Label(value = "XLSX", //
                description = "The Office Open XML format is the file format used by default from Excel 2007 onwards. "
                    + "The maximum number of columns and rows held by a spreadsheet of this format is 16384 and "
                    + "1048576 respectively.") //
            XLSX, //
            @Label(value = "XLS", //
                description = "This is the file format which was used by default up until Excel 2003. "
                    + "The maximum number of columns and rows held by a spreadsheet of this format is 256 and 65536 "
                    + "respectively.") //
            XLS, //
        ;
    }

    private static class SheetExistsPolicyRef implements ParameterReference<SheetExistsPolicy> {
    }

    enum SheetExistsPolicy {
            @Label(value = "Fail",
                description = "Will issue an error during the node's execution (to prevent unintentional overwrite).")
            FAIL, //
            @Label(value = "Overwrite", description = "Will replace any existing sheet.")
            OVERWRITE, //
            @Label(value = "Append",
                description = "Will append the input tables after the last row of the sheets. "
                    + "<b>Note:</b> the last row is chosen according to the rows which already exist in the sheet. "
                    + "A row may appear empty but still exist because it or one of its cells contains styling "
                    + "information or was not removed by Excel after the user cleared it.")
            APPEND, //
        ;
    }

    private static class MissingValuePatternPersistor implements NodeParametersPersistor<Optional<String>> {

        @Override
        public Optional<String> load(final NodeSettingsRO settings) throws InvalidSettingsException {
            if (settings.containsKey(ExcelTableWriterConfig.CFG_REPLACE_MISSINGS)
                && settings.getBoolean(ExcelTableWriterConfig.CFG_REPLACE_MISSINGS)) {
                String pattern = settings.getString(ExcelTableWriterConfig.CFG_MISSING_VALUE_PATTERN, "");
                return Optional.of(pattern);
            }
            return Optional.empty();
        }

        @Override
        public void save(final Optional<String> param, final NodeSettingsWO settings) {
            if (param.isPresent()) {
                settings.addBoolean(ExcelTableWriterConfig.CFG_REPLACE_MISSINGS, true);
                settings.addString(ExcelTableWriterConfig.CFG_MISSING_VALUE_PATTERN, param.get());
            } else {
                settings.addBoolean(ExcelTableWriterConfig.CFG_REPLACE_MISSINGS, false);
                settings.addString(ExcelTableWriterConfig.CFG_MISSING_VALUE_PATTERN, "");
            }
        }

        @Override
        public String[][] getConfigPaths() {
            return new String[][]{{ExcelTableWriterConfig.CFG_REPLACE_MISSINGS},
                {ExcelTableWriterConfig.CFG_MISSING_VALUE_PATTERN}};
        }
    }

    private static class EncryptionRef implements ParameterReference<ExcelEncryptionSettings> {
    }
}
