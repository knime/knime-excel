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

package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

import java.util.List;
import java.util.function.Supplier;

import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileReaderWidget;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileSelection;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.LegacyReaderFileSelectionPersistor;
import org.knime.core.webui.node.dialog.defaultdialog.internal.persistence.PersistArray;
import org.knime.core.webui.node.dialog.defaultdialog.internal.widget.PersistWithin;
import org.knime.core.webui.node.dialog.defaultdialog.widget.Modification;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelEncryptionSettings;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.array.ArrayWidget;
import org.knime.node.parameters.array.ArrayWidget.ElementLayout;
import org.knime.node.parameters.array.PerPortValueProvider;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.layout.Section;
import org.knime.node.parameters.migration.LoadDefaultsForAbsentFields;
import org.knime.node.parameters.migration.Migrate;
import org.knime.node.parameters.persistence.Persist;
import org.knime.node.parameters.persistence.Persistor;
import org.knime.node.parameters.persistence.legacy.LegacyFileWriterWithOverwritePolicyOptions;
import org.knime.node.parameters.persistence.legacy.LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueProvider;
import org.knime.node.parameters.updates.ValueReference;

/**
 * Node parameters for Excel Cell Updater.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline v1.2
 */
@SuppressWarnings("restriction")
@LoadDefaultsForAbsentFields
class ExcelCellUpdaterNodeParameters implements NodeParameters {

    @Section(title = "Input")
    private interface InputSection {
    }

    @Section(title = "Output")
    private interface OutputSection {

    }

    @Section(title = "Update")
    private interface UpdateSection {
    }

    private interface CreateNewFileRef extends ParameterReference<Boolean> {
    }

    interface SrcFileSelectionRef extends ParameterReference<FileSelection> {
    }

    interface EncryptionRef extends ParameterReference<ExcelEncryptionSettings> {
    }

    private interface SheetUpdatesRef extends ParameterReference<SheetUpdate[]> {
    }

    private interface ReplaceMissingsRef extends ParameterReference<Boolean> {
    }

    @Layout(InputSection.class)
    @Widget(title = "Input file", description = """
            Select the Excel file (.xls or .xlsx) to read from and update. The file system can be a local
            file system, a mountpoint, or a file system connected via an input port.
            """)
    @Persistor(SrcFileChooserPersistor.class)
    @Migrate(loadDefaultIfAbsent = true)
    @FileReaderWidget(fileExtensions = {"xls", "xlsx"})
    @ValueReference(SrcFileSelectionRef.class)
    FileSelection m_srcFileSelection = new FileSelection();

    @Layout(InputSection.class)
    @PersistWithin.PersistEmbedded // skip persisting an empty entry, this field only contains one Void field
    ExcelCellUpdaterFileInfoMessage m_fileInfoMessage = new ExcelCellUpdaterFileInfoMessage();

    @Layout(InputSection.class)
    @Widget(title = "Password to protect files", description = """
            Allows you to specify a password to decrypt existing files that have been protected
            with a password via Excel.
            If you specify a password, the output file (or the original, updated file if "Create a new file" is not
            selected) will also be protected with the same password, even if the original input file was not
            password-protected.
            """)
    @Persist(configKey = "authentication_method")
    @ValueReference(EncryptionRef.class)
    @Advanced
    ExcelEncryptionSettings m_encryption = new ExcelEncryptionSettings();

    @Layout(OutputSection.class)
    @Widget(title = "Create a new file", description = """
            If selected, a copy of the original file is created and updated instead of the original file.
            The newly created file must have the same file type as the original file.
            """)
    @Persist(configKey = "create_new_file")
    @ValueReference(CreateNewFileRef.class)
    boolean m_createNewFile;

    @Layout(OutputSection.class)
    @Modification(DestFileModification.class)
    @Persist(configKey = "dest_file_selection")
    @Effect(predicate = CreateNewFileIsTrue.class, type = EffectType.SHOW)
    LegacyFileWriterWithOverwritePolicyOptions m_destFileSelection = new LegacyFileWriterWithOverwritePolicyOptions();

    @Layout(UpdateSection.class)
    @Widget(title = "Input tables",
        description = "For each input table, specify which sheet in the Excel file should be updated.")
    @ArrayWidget(hasFixedSize = true, elementLayout = ElementLayout.VERTICAL_CARD, elementTitle = "Input table")
    @ValueProvider(SheetUpdatesValueProvider.class)
    @PersistArray(SheetUpdate.Persistor.class)
    @ValueReference(SheetUpdatesRef.class)
    SheetUpdate[] m_sheetUpdates = new SheetUpdate[0];

    @Layout(UpdateSection.class)
    @Advanced
    @Widget(title = "Replace missing values by", description = """
            If selected, the node will write the specified value as a string to the address if only
            missing values are found in that row. Otherwise, a blank cell is created.
            """)
    @Persist(configKey = "replace_missings")
    @ValueReference(ReplaceMissingsRef.class)
    boolean m_replaceMissings;

    @Layout(UpdateSection.class)
    @Advanced
    @Widget(title = "Replacement value", description = "The value to write in place of a missing value.")
    @Persist(configKey = "missing_value_pattern")
    @Effect(predicate = ReplaceMissingsIsTrue.class, type = EffectType.SHOW)
    String m_missingValuePattern = "";

    @Layout(UpdateSection.class)
    @Advanced
    @Widget(title = "Evaluate formulas", description = """
            If checked, all formulas in the file will be evaluated after the sheet's updates have been
            performed. This is useful if values that are used in formulas in the sheets have been updated.
            Note: This can cause errors when there are functions not implemented by the Apache POI library.
            """)
    @Persist(configKey = "evaluate_formulas")
    boolean m_evaluateFormulas;

    private static final class SrcFileChooserPersistor extends LegacyReaderFileSelectionPersistor {
        SrcFileChooserPersistor() {
            super("src_file_selection");
        }
    }

    private static final class DestFileModification implements LegacyFileWriterWithOverwritePolicyOptions.Modifier {
        private static final class DestOverwritePolicyChoicesProvider
            extends LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicyChoicesProvider {
            @Override
            protected List<OverwritePolicy> getChoices() {
                return List.of(OverwritePolicy.fail, OverwritePolicy.overwrite);
            }
        }

        @Override
        public void modify(final Modification.WidgetGroupModifier group) {
            var fileSelection = findFileSelection(group);
            fileSelection.modifyAnnotation(Widget.class) //
                .withProperty("title", "Output file") //
                .withProperty("description", """
                        Provide the path to the new file that should be created as a copy of the original file.
                        The output file must have the same file type (.xls or .xlsx) as the input file.
                        """) //
                .modify();
            restrictOverwritePolicyOptions(group, DestOverwritePolicyChoicesProvider.class);
        }
    }

    private static final class SheetUpdatesValueProvider extends PerPortValueProvider<SheetUpdate> {

        SheetUpdatesValueProvider() {
            super(ExcelCellUpdaterNodeFactory.SHEET_GRP_ID, PortGroupSide.INPUT);
        }

        @Override
        protected Supplier<SheetUpdate[]> supplier(final StateProviderInitializer initializer) {
            return initializer.getValueSupplier(SheetUpdatesRef.class);
        }

        @Override
        protected SheetUpdate[] newArray(final int size) {
            return new SheetUpdate[size];
        }

        @Override
        protected SheetUpdate newInstance(final int index) {
            return new SheetUpdate(index);
        }
    }

    private static final class CreateNewFileIsTrue implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getBoolean(CreateNewFileRef.class).isTrue();
        }
    }

    private static final class ReplaceMissingsIsTrue implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getBoolean(ReplaceMissingsRef.class).isTrue();
        }
    }

}
