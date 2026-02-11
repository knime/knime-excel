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

import java.io.IOException;
import java.util.Collection;
import java.util.List;
import java.util.stream.Stream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileReaderParameters.HowToCombineColumnsOption;
import org.knime.base.node.io.filehandling.webui.testing.reader2.TransformationParametersUpdatesTest;
import org.knime.base.node.io.filehandling.webui.reader2.ReaderSpecific;
import org.knime.base.node.io.filehandling.webui.reader2.TransformationParameters;
import org.knime.core.data.DataType;
import org.knime.core.data.def.BooleanCell;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;
import org.knime.core.node.workflow.CredentialsProvider;
import org.knime.core.node.workflow.NodeContext;
import org.knime.core.util.Pair;
import org.knime.core.webui.node.dialog.defaultdialog.NodeParametersInputImpl;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.testing.MountPointFileSystemAccessMock;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.node.table.reader.ProductionPathProvider;
import org.knime.filehandling.core.node.table.reader.config.tablespec.ProductionPathSerializer;
import org.knime.node.parameters.NodeParametersInput;

/**
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
final class ExcelTableReaderTransformationParametersStateProvidersTest
    extends TransformationParametersUpdatesTest<ExcelTableReaderNodeParameters, ExcelCell.KNIMECellType> {

    @BeforeEach
    @Override
    protected void setUpSettingsAndFile() {
        MountPointFileSystemAccessMock.enabled = true;
        NodeContext.pushContext(m_wfm.getNodeContainer(m_wfm.createAndAddNode(new ExcelTableReaderNodeFactory())));
        super.setUpSettingsAndFile();
    }

    @AfterEach
    void teardown() {
        NodeContext.removeLastContext();
        MountPointFileSystemAccessMock.enabled = false;
    }

    @Override
    protected NodeParametersInput createNodeParametersInput() {
        return NodeParametersInputImpl.createDefaultNodeSettingsContext(new PortType[0], new PortObjectSpec[0],
                null, CredentialsProvider.EMPTY_CREDENTIALS_PROVIDER);
    }

    // Trigger paths for Excel reader-specific parameters that affect table spec (ConfigID)
    // Note: Encryption settings also affect spec (encrypted files require credentials to read)
    static final List<List<String>> TRIGGER_PATHS = List.of( //
        // SelectSheetParameters (from saveConfigIDSettingsTab -> saveSettingsTab)
        List.of("excelTableReaderParameters", "selectSheetParams", "sheetSelection"), //
        List.of("excelTableReaderParameters", "selectSheetParams", "sheetName"), //
        List.of("excelTableReaderParameters", "selectSheetParams", "sheetIndex"), //

        // ReadAreaParameters (from saveConfigIDSettingsTab -> saveSettingsTab)
        List.of("excelTableReaderParameters", "readAreaParams", "sheetArea"), //
        List.of("excelTableReaderParameters", "readAreaParams", "readFromColumn"), //
        List.of("excelTableReaderParameters", "readAreaParams", "readToColumn"), //
        List.of("excelTableReaderParameters", "readAreaParams", "readFromRow"), //
        List.of("excelTableReaderParameters", "readAreaParams", "readToRow"), //

        // ColumnNameParameters (from saveConfigIDSettingsTab -> saveSettingsTab)
        List.of("excelTableReaderParameters", "columnNameParams", "columnNameMode"), //
        List.of("excelTableReaderParameters", "columnNameParams", "columnNamesRowNumber"), //
        List.of("excelTableReaderParameters", "columnNameParams", "emptyColHeaderPrefix"), //

        // RowIdParameters (from saveConfigIDSettingsTab -> saveSettingsTab)
        List.of("excelTableReaderParameters", "rowIdParams", "rowIdMode"), //
        List.of("excelTableReaderParameters", "rowIdParams", "rowIdColumn"), //

        // SkipParameters (from saveConfigIDAdvancedSettingsTab - ONLY these 3, NOT skipEmptyRows)
        List.of("excelTableReaderParameters", "skipParams", "skipEmptyColumns"), //
        List.of("excelTableReaderParameters", "skipParams", "skipHiddenColumns"), //
        List.of("excelTableReaderParameters", "skipParams", "skipHiddenRows"), //

        // SchemaDetectionParameters (from saveConfigIDAdvancedSettingsTab - ONLY these 2, NOT failOnDifferingSpecs)
        List.of("excelTableReaderParameters", "schemaDetectionParams", "limitDataRowsScanned"), //
        List.of("excelTableReaderParameters", "schemaDetectionParams", "maxDataRowsScanned"), //

        // ValuesParameters (from saveConfigIDAdvancedSettingsTab)
        List.of("excelTableReaderParameters", "valuesParams", "use15DigitsPrecision"), //
        List.of("excelTableReaderParameters", "valuesParams", "replaceEmptyStringsWithMissingValues"), //
        List.of("excelTableReaderParameters", "valuesParams", "reevaluateFormulas"), //
        List.of("excelTableReaderParameters", "valuesParams", "formulaErrorHandling"), //

        // EncryptionParameters (required for reading encrypted files - affects spec detection)
        List.of("excelTableReaderParameters", "encryptionParams", "encryptionMode"), //
        List.of("excelTableReaderParameters", "encryptionParams", "credentials") //
    );

    @Override
    protected void setSourcePath(final ExcelTableReaderNodeParameters settings, final FSLocation fsLocation) {
        settings.m_excelTableReaderParameters.m_multiFileSelectionParams.m_source.m_path = fsLocation;
    }

    @Override
    protected void setHowToCombineColumns(final ExcelTableReaderNodeParameters settings,
        final HowToCombineColumnsOption howToCombineColumns) {
        settings.m_excelTableReaderParameters.m_multiFileReaderParams.m_howToCombineColumns = howToCombineColumns;
    }

    @Override
    protected TransformationParameters<ExcelCell.KNIMECellType>
        getTransformationSettings(final ExcelTableReaderNodeParameters params) {
        return params.m_transformationParameters;
    }

    @Override
    protected void writeFileWithIntegerAndStringColumn(final String filePath) throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Sheet1");
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("intCol");
            headerRow.createCell(1).setCellValue("stringCol");
            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(1);
            row1.createCell(1).setCellValue("value1");
            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue(2);
            row2.createCell(1).setCellValue("value2");

            try (var outputStream = new java.io.FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        }
    }

    @Override
    protected ProductionPathProvider<ExcelCell.KNIMECellType> getProductionPathProvider() {
        return ExcelTableReaderSpecific.PRODUCTION_PATH_PROVIDER;
    }

    @Override
    protected Pair<DataType, Collection<IntOrString>> getUnreachableType() {
        return new Pair<>(BooleanCell.TYPE, List.of(IntOrString.INT));
    }

    @Override
    protected ExcelCell.KNIMECellType getIntType() {
        return ExcelCell.KNIMECellType.INT;
    }

    @Override
    protected ExcelCell.KNIMECellType getStringType() {
        return ExcelCell.KNIMECellType.STRING;
    }

    @Override
    protected ExcelCell.KNIMECellType getDoubleType() {
        return ExcelCell.KNIMECellType.DOUBLE;
    }

    @Override
    protected ProductionPathSerializer getProductionPathSerializer() {
        return new ExcelTableReaderTransformationParameters().getProductionPathSerializer();
    }

    @Override
    protected List<String> getPathToTransformationSettings() {
        return List.of("transformationParameters");
    }

    @Override
    protected ExcelTableReaderNodeParameters constructNewSettings() {
        return new ExcelTableReaderNodeParameters();
    }

    @Override
    protected String getFileName() {
        return "test.xlsx";
    }

    @Override
    protected ReaderSpecific.ExternalDataTypeSerializer<ExcelCell.KNIMECellType> getExternalDataTypeSerializer() {
        return new ExcelTableReaderTransformationParameters();
    }

    @ParameterizedTest
    @MethodSource("getDependencyTriggerReferences")
    void testTableSpecSettingsProvider(final List<String> triggerPath) throws IOException {
        testTableSpecSettingsProvider(sim -> sim.simulateValueChange(triggerPath));
    }

    static Stream<Arguments> getDependencyTriggerReferences() {
        return TRIGGER_PATHS.stream().map(Arguments::of);
    }

}
