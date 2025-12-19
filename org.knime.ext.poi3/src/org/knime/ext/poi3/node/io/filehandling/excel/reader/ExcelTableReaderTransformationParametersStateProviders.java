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
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.util.function.Supplier;

import org.knime.base.node.io.filehandling.webui.reader2.TransformationParametersStateProviders;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ColumnNameParameters.EmptyColHeaderSuffixRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderNodeParameters.ExcelTableReaderParametersRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderSpecific.ConfigAndReader;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderSpecific.ProductionPathProviderAndTypeHierarchy;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.SkipParameters.SkipEmptyColumnsRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.SkipParameters.SkipHiddenColumnsRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.SkipParameters.SkipHiddenRowsRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ValuesParameters.FormulaErrorHandlingRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ValuesParameters.ReevaluateFormulasRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ValuesParameters.ReplaceEmptyStringsWithMissingValuesRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ValuesParameters.Use15DigitsPrecisionRef;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;

/**
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
final class ExcelTableReaderTransformationParametersStateProviders {

    static final class TypedReaderTableSpecsProvider
        extends TransformationParametersStateProviders.TypedReaderTableSpecsProvider<//
            ExcelTableReaderConfig, KNIMECellType, ExcelMultiTableReadConfig>
        implements ConfigAndReader {

        private Supplier<ExcelTableReaderParameters> m_readerNodeParamsSupplier;

        @Override
        protected void applyParametersToConfig(final ExcelMultiTableReadConfig config) {
            m_readerNodeParamsSupplier.get().saveToConfig(config);
    }

        @Override
        public void init(final StateProviderInitializer initializer) {
            super.init(initializer);
            m_readerNodeParamsSupplier = initializer.getValueSupplier(ExcelTableReaderParametersRef.class);
        }

        interface Dependent
            extends TransformationParametersStateProviders.TypedReaderTableSpecsProvider.Dependent<KNIMECellType> {
            @SuppressWarnings("unchecked")
            @Override
            default Class<TypedReaderTableSpecsProvider> getTypedReaderTableSpecsProvider() {
                return TypedReaderTableSpecsProvider.class;
            }

            @Override
            default void initConfigIdTriggers(final StateProviderInitializer initializer) {
                // Register triggers for all reader-specific parameters that affect table spec detection

                initializer.computeOnValueChange(EncryptionParameters.EncryptionModeRef.class);
                initializer.computeOnValueChange(EncryptionParameters.CredentialsRef.class);

                // Sheet selection parameters affect table spec
                initializer.computeOnValueChange(SelectSheetParameters.SheetSelectionRef.class);
                initializer.computeOnValueChange(SelectSheetParameters.SheetNameRef.class);
                initializer.computeOnValueChange(SelectSheetParameters.SheetIndexRef.class);

                // Read area parameters affect table spec
                initializer.computeOnValueChange(ReadAreaParameters.SheetAreaRef.class);
                initializer.computeOnValueChange(ReadAreaParameters.ReadFromColumnRef.class);
                initializer.computeOnValueChange(ReadAreaParameters.ReadToColumnRef.class);
                initializer.computeOnValueChange(ReadAreaParameters.ReadFromRowRef.class);
                initializer.computeOnValueChange(ReadAreaParameters.ReadToRowRef.class);

                // Column name parameters affect table spec
                initializer.computeOnValueChange(ColumnNameParameters.ColumnNameModeRef.class);
                initializer.computeOnValueChange(ColumnNameParameters.ColumnNamesRowNumberRef.class);
                initializer.computeOnValueChange(ColumnNameParameters.EmptyColHeaderPrefixRef.class);
                initializer.computeOnValueChange(EmptyColHeaderSuffixRef.class);

                // RowID parameters affect table spec
                initializer.computeOnValueChange(RowIdParameters.RowIdModeRef.class);
                initializer.computeOnValueChange(RowIdParameters.RowIdColumnRef.class);

                // Schema detection parameters affect table spec
                initializer.computeOnValueChange(SchemaDetectionParameters.LimitDataRowsScannedRef.class);
                initializer.computeOnValueChange(SchemaDetectionParameters.MaxDataRowsScannedRef.class);
                initializer.computeOnValueChange(SchemaDetectionParameters.FailOnDifferingSpecsRef.class);

                // Skip parameters affect table spec (ONLY these 3, NOT skipEmptyRows)
                initializer.computeOnValueChange(SkipEmptyColumnsRef.class);
                initializer.computeOnValueChange(SkipHiddenColumnsRef.class);
                initializer.computeOnValueChange(SkipHiddenRowsRef.class);

                // Values parameters affect table spec
                initializer.computeOnValueChange(Use15DigitsPrecisionRef.class);
                initializer.computeOnValueChange(ReplaceEmptyStringsWithMissingValuesRef.class);
                initializer.computeOnValueChange(ReevaluateFormulasRef.class);
                initializer.computeOnValueChange(FormulaErrorHandlingRef.class);
            }
        }
    }

    static final class TableSpecSettingsProvider
        extends TransformationParametersStateProviders.TableSpecSettingsProvider<KNIMECellType>
        implements TypedReaderTableSpecsProvider.Dependent, KNIMECellTypeSerializer {
    }

    static final class TransformationElementSettingsProvider
        extends TransformationParametersStateProviders.TransformationElementSettingsProvider<KNIMECellType> implements
        ProductionPathProviderAndTypeHierarchy, TypedReaderTableSpecsProvider.Dependent, KNIMECellTypeSerializer {

        @Override
        protected boolean hasMultipleFileHandling() {
            return true;
        }
    }

    static final class TypeChoicesProvider
        extends TransformationParametersStateProviders.TypeChoicesProvider<KNIMECellType>
        implements ProductionPathProviderAndTypeHierarchy, TypedReaderTableSpecsProvider.Dependent {
    }

    static final class TransformationSettingsWidgetModification
        extends TransformationParametersStateProviders.TransformationSettingsWidgetModification<KNIMECellType> {

        @Override
        protected Class<TableSpecSettingsProvider> getSpecsValueProvider() {
            return TableSpecSettingsProvider.class;
        }

        @Override
        protected Class<TypeChoicesProvider> getTypeChoicesProvider() {
            return TypeChoicesProvider.class;
        }

        @Override
        protected Class<TransformationElementSettingsProvider> getTransformationSettingsValueProvider() {
            return TransformationElementSettingsProvider.class;
        }

    }

    private ExcelTableReaderTransformationParametersStateProviders() {
        // Not intended to be initialized
    }

}
