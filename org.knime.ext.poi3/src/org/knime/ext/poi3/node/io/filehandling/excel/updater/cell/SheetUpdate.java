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
import java.util.Optional;
import java.util.function.Supplier;

import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.StringValue;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.webui.node.dialog.defaultdialog.internal.persistence.ArrayPersistor;
import org.knime.core.webui.node.dialog.defaultdialog.internal.persistence.ElementFieldPersistor;
import org.knime.core.webui.node.dialog.defaultdialog.internal.persistence.PersistArrayElement;
import org.knime.core.webui.node.dialog.defaultdialog.util.updates.StateComputationFailureException;
import org.knime.core.webui.node.dialog.defaultdialog.widget.Modification;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelFileContentInfoStateProvider.ExcelFileInfo;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.StateProvider;
import org.knime.node.parameters.updates.ValueProvider;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.updates.legacy.AutoGuessValueProvider;
import org.knime.node.parameters.widget.choices.ChoicesProvider;
import org.knime.node.parameters.widget.choices.StringChoicesProvider;
import org.knime.node.parameters.widget.choices.TypedStringChoice;
import org.knime.node.parameters.widget.choices.util.CompatibleColumnsProvider.StringColumnsProvider;

import static org.knime.ext.poi3.node.io.filehandling.excel.updater.cell.ExcelCellUpdaterNodeFactory.SHEET_GRP_ID;

/**
 * Parameters for a single sheet update entry: holds the target sheet name for one input table port.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 */
class SheetUpdate implements NodeParameters {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(SheetUpdate.class);

    @SuppressWarnings("unused") // needed by framework
    SheetUpdate() {
    }

    SheetUpdate(final int index) {
        this(index, null, null);
    }

    SheetUpdate(final int index, final String sheetName, final String coordinateColumnName) {
        m_index = index;
        m_sheetName = sheetName;
        m_coordinateColumnName = coordinateColumnName;
    }

    private interface IndexRef extends ParameterReference<Integer> {
    }

    private interface SheetNameRef extends ParameterReference<String> {
    }

    static final class CoordinateColumnRef implements ParameterReference<String>, Modification.Reference {
    }

    @ValueReference(IndexRef.class)
    @PersistArrayElement(Persistor.IndexPersistor.class)
    int m_index;

    @Widget(title = "Excel sheet",
        description = "Specify which sheet should be updated by the corresponding input table.")
    @ChoicesProvider(SheetNamesChoicesProvider.class)
    @ValueProvider(SheetNameDefaultProvider.class)
    @ValueReference(SheetNameRef.class)
    @PersistArrayElement(Persistor.SheetNamePersistor.class)
    String m_sheetName; // the node model checks for null values during configure, therefore not initializing with ""

    @Widget(title = "Based on address column",
        description = "Select the name of the column which contains the addresses of the cells"
            + " which should be updated.")
    @ChoicesProvider(CoordinateColumnChoicesProvider.class)
    @ValueProvider(CoordinateColumnDefaultProvider.class)
    @Modification.WidgetReference(CoordinateColumnRef.class)
    @ValueReference(CoordinateColumnRef.class)
    @PersistArrayElement(Persistor.CoordinateColumnNameElementPersistor.class)
    String m_coordinateColumnName;

    private static class SheetNamesChoicesProvider implements StringChoicesProvider {

        Supplier<ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProvider.StateProviderInitializer initializer) {
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelCellUpdaterFileContentStateProvider.class);
        }

        @Override
        public List<String> choices(final NodeParametersInput context) {
            return m_excelFileInfo.get().getSheetNames();
        }
    }

    private static class SheetNameDefaultProvider extends AutoGuessValueProvider<String> {

        SheetNameDefaultProvider() {
            super(SheetNameRef.class);
        }

        Supplier<ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProviderInitializer initializer) {
            super.init(initializer);
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelCellUpdaterFileContentStateProvider.class);
        }

        @Override
        protected boolean isEmpty(final String value) {
            return value == null || value.isEmpty();
        }

        @Override
        protected String autoGuessValue(final NodeParametersInput parametersInput)
            throws StateComputationFailureException {
            return Optional.of(m_excelFileInfo.get()).map(ExcelFileInfo::getSheetNames)
                .flatMap(names -> names.stream().findFirst()).orElseThrow(StateComputationFailureException::new);
        }
    }

    private static class CoordinateColumnChoicesProvider extends StringColumnsProvider {

        Supplier<Integer> m_index;

        @Override
        public void init(StateProviderInitializer initializer) {
            super.init(initializer);
            m_index = initializer.computeFromValueSupplier(IndexRef.class);
        }

        @Override
        public List<TypedStringChoice> computeState(NodeParametersInput context) {
            try {
                return super.computeState(context);
            } catch (IndexOutOfBoundsException e) { // NOSONAR
                // this can happen when the port index is not available anymore when a port is removed
                return List.of();
            }
        }

        @Override
        public int getInputTableIndex(NodeParametersInput parametersInput) {
            var portIndices = parametersInput.getPortsConfiguration().getInputPortLocation().get(SHEET_GRP_ID);
            return portIndices[m_index.get()];
        }
    }

    private static class CoordinateColumnDefaultProvider extends AutoGuessValueProvider<String> {

        Supplier<Integer> m_index;

        @Override
        public void init(StateProviderInitializer initializer) {
            super.init(initializer);
            m_index = initializer.computeFromValueSupplier(IndexRef.class);
        }

        CoordinateColumnDefaultProvider() {
            super(CoordinateColumnRef.class);
        }

        @Override
        protected boolean isEmpty(final String value) {
            return value == null || value.isEmpty();
        }

        @Override
        protected String autoGuessValue(final NodeParametersInput parametersInput)
            throws StateComputationFailureException {
            final var portIndices = parametersInput.getPortsConfiguration().getInputPortLocation().get(SHEET_GRP_ID);
            final int portIndex;
            try {
                portIndex = portIndices[m_index.get()];
            } catch (IndexOutOfBoundsException e) {
                LOGGER.warn("Could not determine input table for coordinate column suggestions.", e);
                throw new StateComputationFailureException();
            }
            return parametersInput.getInTableSpec(portIndex) //
                .orElseThrow(StateComputationFailureException::new) //
                .stream() //
                .filter(col -> col.getType().isCompatible(StringValue.class)) //
                .map(DataColumnSpec::getName) //
                .findFirst() //
                .orElseThrow(StateComputationFailureException::new);
        }
    }

    static final class Persistor implements ArrayPersistor<Integer, Persistor.SheetUpdateSaveDTO> {

        static final String SHEET_NAMES = "sheet_names";

        static final String COORDINATE_COLUMN_NAMES = "coordinate_column_names";

        @Override
        public int getArrayLength(final NodeSettingsRO nodeSettings) throws InvalidSettingsException {
            // be tolerant against flow variables setting only one array or two arrays with differing lengths
            return Math.min( //
                nodeSettings.getStringArray(SHEET_NAMES).length,
                nodeSettings.getStringArray(COORDINATE_COLUMN_NAMES).length);
        }

        @Override
        public Integer createElementLoadContext(final int index) {
            return index;
        }

        @SuppressWarnings("ClassEscapesDefinedScope")
        @Override
        public SheetUpdateSaveDTO createElementSaveDTO(final int index) {
            return new SheetUpdateSaveDTO(index);
        }

        @SuppressWarnings("ClassEscapesDefinedScope")
        @Override
        public void save(final List<SheetUpdateSaveDTO> savedElements, final NodeSettingsWO nodeSettings) {
            final var sheetNames = new String[savedElements.size()];
            final var coordinateColumnNames = new String[savedElements.size()];
            for (final var dto : savedElements) {
                sheetNames[dto.m_index] = dto.m_sheetName;
                coordinateColumnNames[dto.m_index] = dto.m_coordinateColumnName;
            }
            nodeSettings.addStringArray(SHEET_NAMES, sheetNames);
            nodeSettings.addStringArray(COORDINATE_COLUMN_NAMES, coordinateColumnNames);
        }

        private static final class SheetUpdateSaveDTO {
            final int m_index;

            String m_sheetName = "";

            String m_coordinateColumnName = "";

            SheetUpdateSaveDTO(final int index) {
                m_index = index;
            }
        }

        private static final class IndexPersistor
            implements ElementFieldPersistor<Integer, Integer, SheetUpdateSaveDTO> {

            @Override
            public Integer load(final NodeSettingsRO nodeSettings, final Integer loadContext) {
                return loadContext;
            }

            @Override
            public void save(final Integer param, final SheetUpdateSaveDTO saveDTO) {
                // nothing to do
            }

            @Override
            public String[][] getConfigPaths() {
                return new String[0][];
            }
        }

        private static final class SheetNamePersistor
            implements ElementFieldPersistor<String, Integer, SheetUpdateSaveDTO> {

            @Override
            public String load(final NodeSettingsRO nodeSettings, final Integer loadContext)
                throws InvalidSettingsException {
                return nodeSettings.getStringArray(Persistor.SHEET_NAMES)[loadContext];
            }

            @Override
            public void save(final String param, final SheetUpdateSaveDTO saveDTO) {
                saveDTO.m_sheetName = param;
            }

            @Override
            public String[][] getConfigPaths() {
                return new String[][]{{Persistor.SHEET_NAMES}};
            }
        }

        private static final class CoordinateColumnNameElementPersistor
            implements ElementFieldPersistor<String, Integer, SheetUpdateSaveDTO> {

            @Override
            public String load(final NodeSettingsRO nodeSettings, final Integer loadContext)
                throws InvalidSettingsException {
                return nodeSettings.getStringArray(Persistor.COORDINATE_COLUMN_NAMES)[loadContext];
            }

            @Override
            public void save(final String param, final SheetUpdateSaveDTO saveDTO) {
                saveDTO.m_coordinateColumnName = param;
            }

            @Override
            public String[][] getConfigPaths() {
                return new String[][]{{Persistor.COORDINATE_COLUMN_NAMES}};
            }
        }
    }
}
