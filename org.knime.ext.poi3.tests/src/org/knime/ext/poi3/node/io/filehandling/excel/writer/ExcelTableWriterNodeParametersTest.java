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
 *   Dec 15, 2025: created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.List;

import org.eclipse.core.runtime.FileLocator;
import org.junit.jupiter.api.Test;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.util.FileUtil;
import org.knime.core.webui.node.dialog.SettingsType;
import org.knime.core.webui.node.dialog.defaultdialog.NodeParametersUtil;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileSelection;
import org.knime.ext.poi3.Fixtures;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelEncryptionSettings;
import org.knime.filehandling.core.connections.FSCategory;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.node.parameters.persistence.legacy.LegacyFileWriterWithOverwritePolicyOptions;
import org.knime.node.parameters.widget.choices.StringChoice;
import org.knime.node.parameters.widget.credentials.Credentials;
import org.knime.testing.node.dialog.DefaultNodeSettingsSnapshotTest;
import org.knime.testing.node.dialog.SnapshotTestConfiguration;
import org.knime.testing.node.dialog.updates.DialogUpdateSimulator;

/**
 * Snapshot test for {@link ExcelTableWriterNodeParameters}.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline
 */
@SuppressWarnings("restriction")
final class ExcelTableWriterNodeParametersTest extends DefaultNodeSettingsSnapshotTest {

    ExcelTableWriterNodeParametersTest() {
        super(getConfig());
    }

    private static SnapshotTestConfiguration getConfig() {
        return SnapshotTestConfiguration.builder() //
            .withInputPortObjectSpecs(createInputPortSpecs()) //
            .testJsonFormsForModel(ExcelTableWriterNodeParameters.class) //
            .testJsonFormsWithInstance(SettingsType.MODEL, () -> readSettings()) //
            .testNodeSettingsStructure(() -> readSettings()) //
            .build();
    }

    private static ExcelTableWriterNodeParameters readSettings() {
        try {
            var path = getSnapshotPath(ExcelTableWriterNodeParameters.class).getParent().resolve("node_settings")
                .resolve("ExcelTableWriterNodeParameters.xml");
            try (var fis = new FileInputStream(path.toFile())) {
                var nodeSettings = NodeSettings.loadFromXML(fis);
                return NodeParametersUtil.loadSettings(nodeSettings.getNodeSettings(SettingsType.MODEL.getConfigKey()),
                    ExcelTableWriterNodeParameters.class);
            }
        } catch (IOException | InvalidSettingsException e) {
            throw new IllegalStateException(e);
        }
    }

    private static PortObjectSpec[] createInputPortSpecs() {
        return new PortObjectSpec[]{createDefaultTestTableSpec(), createDefaultTestTableSpec(),
            createDefaultTestTableSpec()};
    }

    private static DataTableSpec createDefaultTestTableSpec() {
        return new DataTableSpec(new DataColumnSpecCreator("Column1", StringCell.TYPE).createSpec(),
            new DataColumnSpecCreator("Column2", StringCell.TYPE).createSpec());
    }

    @Test
    @SuppressWarnings("static-method")
    void testExtensionUpdateOnExcelFormatChange() {
        // given
        var settings = new ExcelTableWriterNodeParameters();
        var outputFile = new FSLocation(FSCategory.LOCAL, "workbook.xlsx");
        var outputFileSelection = new FileSelection(outputFile);
        settings.m_outputFile = new LegacyFileWriterWithOverwritePolicyOptions(outputFileSelection, true,
            LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.overwrite);
        var simulator = new DialogUpdateSimulator(settings, null);

        // when
        settings.m_excelFormat = ExcelTableWriterNodeParameters.ExcelFormat.XLS;
        var result = simulator.simulateValueChange("excelFormat");

        // then
        assertThat(result.getValueUpdateAt("outputFile", "file"))
            .isEqualTo(new FileSelection(new FSLocation(FSCategory.LOCAL, "workbook.xls")));
        assertThat(result.getUiStateUpdateAt(List.of("outputFile", "file"), "fileExtension")).isEqualTo("xls");
    }

    @Test
    @SuppressWarnings("static-method")
    void testUnknownExtensionIsNotUpdatedOnExcelFormatChange() {
        // given
        var settings = new ExcelTableWriterNodeParameters();
        var outputFile = new FSLocation(FSCategory.LOCAL, "workbook.unknown");
        var outputFileSelection = new FileSelection(outputFile);
        settings.m_outputFile = new LegacyFileWriterWithOverwritePolicyOptions(outputFileSelection, true,
            LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.overwrite);
        var simulator = new DialogUpdateSimulator(settings, null);

        // when
        settings.m_excelFormat = ExcelTableWriterNodeParameters.ExcelFormat.XLS;
        var result = simulator.simulateValueChange("excelFormat");

        // then
        assertThat(result.hasNoValueUpdateAt("outputFile", "file")).isTrue();
        assertThat(result.getUiStateUpdateAt(List.of("outputFile", "file"), "fileExtension")).isEqualTo("xls");
    }

    @Test
    @SuppressWarnings("static-method")
    void testSheetNamesAreLoadedFromUnencryptedExcelFile() throws IOException, URISyntaxException {
        // given
        var settings = new ExcelTableWriterNodeParameters();
        var outputFile = new FSLocation(FSCategory.LOCAL, getTestFilePath(Fixtures.XLSX));
        var outputFileSelection = new FileSelection(outputFile);
        settings.m_outputFile = new LegacyFileWriterWithOverwritePolicyOptions(outputFileSelection, true,
            LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.overwrite);
        var simulator = new DialogUpdateSimulator(settings, null);

        // when
        var result = simulator.simulateValueChange("outputFile", "file");

        // then
        assertThat(
            result.getUiStateUpdateInArrayAt(List.of(List.of("sheetNames"), List.of("sheetName")), "possibleValues"))
                .isEqualTo(List.of(StringChoice.fromId("knime"), StringChoice.fromId("knime2")));
    }

    @Test
    @SuppressWarnings("static-method")
    void testSheetNamesAreLoadedFromPasswordProtectedExcelFile() throws IOException, URISyntaxException {
        // given
        var settings = new ExcelTableWriterNodeParameters();
        var outputFile = new FSLocation(FSCategory.LOCAL, getTestFilePath(Fixtures.XLSX_ENC));
        var outputFileSelection = new FileSelection(outputFile);
        settings.m_outputFile = new LegacyFileWriterWithOverwritePolicyOptions(outputFileSelection, true,
            LegacyFileWriterWithOverwritePolicyOptions.OverwritePolicy.overwrite);
        settings.m_encryption = new ExcelEncryptionSettings(ExcelEncryptionSettings.AuthType.PWD, null,
            new Credentials("", Fixtures.TEST_PW));
        var simulator = new DialogUpdateSimulator(settings, null);

        // when
        var result = simulator.simulateValueChange("outputFile", "file");

        // then
        assertThat(
            result.getUiStateUpdateInArrayAt(List.of(List.of("sheetNames"), List.of("sheetName")), "possibleValues"))
                .isEqualTo(List.of(StringChoice.fromId("knime"), StringChoice.fromId("knime2")));
    }

    private String getTestFilePath(final String path) throws IOException, URISyntaxException {
        var url = FileLocator.toFileURL(getClass().getResource(path).toURI().toURL());
        return FileUtil.getFileFromURL(url).toPath().toString();
    }

}
