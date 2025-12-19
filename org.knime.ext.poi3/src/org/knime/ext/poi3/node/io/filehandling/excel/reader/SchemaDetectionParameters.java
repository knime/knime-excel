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

import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionParameters.MultiFileSelectionIsActive;
import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderValidation.MaxNumberOfExcelRows;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.Inside;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.number.NumberInputWidget;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation.MinValidation.IsPositiveIntegerValidation;

/**
 * Parameters for limiting the number of scanned rows in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(SchemaDetectionParameters.SchemaDetectionLayout.class)
class SchemaDetectionParameters implements NodeParameters {

    @Inside(ReaderLayout.ColumnAndDataTypeDetection.class)
    interface SchemaDetectionLayout {
    }

    static final class LimitDataRowsScannedRef implements ParameterReference<Boolean> {
    }

    static final class MaxDataRowsScannedRef implements ParameterReference<Long> {
    }

    static final class FailOnDifferingSpecsRef implements ParameterReference<Boolean> {
    }

    private static final class IsLimitEnabled implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getBoolean(LimitDataRowsScannedRef.class).isTrue();
        }
    }

    @Widget(title = "Limit scanned rows",
        description = "If you limit the number of scanned rows, only the specified number of input rows are used "
            + "to analyze the file (i.e. to determine the column types). This option is recommended for long files "
            + "where the first n rows are representative for the whole file.")
    @ValueReference(LimitDataRowsScannedRef.class)
    boolean m_limitDataRowsScanned = true;

    @Widget(title = "Maximum number of scanned rows",
        description = "The maximum number of rows to scan for determining column types.")
    @NumberInputWidget(minValidation = IsPositiveIntegerValidation.class, maxValidation = MaxNumberOfExcelRows.class)
    @ValueReference(MaxDataRowsScannedRef.class)
    @Effect(predicate = IsLimitEnabled.class, type = EffectType.SHOW)
    long m_maxDataRowsScanned = 10000;

    @Widget(title = "Fail if schemas differ between multiple files",
        description = "If checked, the node will fail if multiple files are read via the \"Files in folder\" option "
            + "and not all files have the same table structure, i.e. the same columns.")
    @ValueReference(FailOnDifferingSpecsRef.class)
    @Effect(predicate = MultiFileSelectionIsActive.class, type = EffectType.SHOW)
    boolean m_failOnDifferingSpecs = true;

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var tableReadConfig = config.getTableReadConfig();

        tableReadConfig.setLimitRowsForSpec(m_limitDataRowsScanned);
        if (m_limitDataRowsScanned) {
            tableReadConfig.setMaxRowsForSpec(m_maxDataRowsScanned);
        }

        config.setFailOnDifferingSpecs(m_failOnDifferingSpecs);
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var tableReadConfig = config.getTableReadConfig();

        m_limitDataRowsScanned = tableReadConfig.limitRowsForSpec();
        m_maxDataRowsScanned = tableReadConfig.getMaxRowsForSpec();
        m_failOnDifferingSpecs = config.failOnDifferingSpecs();
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // no further validation needed
    }

}

