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
 *   Feb 25, 2026: created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.List;

import org.eclipse.core.runtime.FileLocator;
import org.junit.jupiter.api.Test;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.util.FileUtil;
import org.knime.core.webui.node.dialog.SettingsType;
import org.knime.core.webui.node.dialog.defaultdialog.NodeParametersUtil;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.FileSelection;
import org.knime.ext.poi3.Fixtures;
import org.knime.ext.poi3.node.io.filehandling.excel.ExcelEncryptionSettings;
import org.knime.filehandling.core.connections.FSCategory;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.node.parameters.widget.choices.StringChoice;
import org.knime.node.parameters.widget.credentials.Credentials;
import org.knime.testing.node.dialog.DefaultNodeSettingsSnapshotTest;
import org.knime.testing.node.dialog.SnapshotTestConfiguration;
import org.knime.testing.node.dialog.updates.DialogUpdateSimulator;

/**
 * Snapshot test for {@link ExcelCellUpdaterNodeParameters}.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline
 */
@SuppressWarnings("restriction")
final class ExcelCellUpdaterNodeParametersTest extends DefaultNodeSettingsSnapshotTest {

    ExcelCellUpdaterNodeParametersTest() {
        super(getConfig());
    }

    private static SnapshotTestConfiguration getConfig() {
        return SnapshotTestConfiguration.builder() //
            .withPortsConfiguration(createPortConfig()) //
            .withInputPortObjectSpecs(createInputPortSpecs()) //
            .testJsonFormsForModel(ExcelCellUpdaterNodeParameters.class) //
            .testJsonFormsWithInstance(SettingsType.MODEL, () -> readSettings()) //
            .testNodeSettingsStructure(() -> readSettings()) //
            .build();
    }

    private static ExcelCellUpdaterNodeParameters readSettings() {
        try {
            var path = getSnapshotPath(ExcelCellUpdaterNodeParameters.class).getParent().resolve("node_settings")
                .resolve("ExcelCellUpdaterNodeParameters.xml");
            try (var fis = new FileInputStream(path.toFile())) {
                var nodeSettings = NodeSettings.loadFromXML(fis);
                return NodeParametersUtil.loadSettings(nodeSettings.getNodeSettings(SettingsType.MODEL.getConfigKey()),
                    ExcelCellUpdaterNodeParameters.class);
            }
        } catch (IOException | InvalidSettingsException e) {
            throw new IllegalStateException(e);
        }
    }

    private static PortsConfiguration createPortConfig() {
        // Two data table input ports in the example snapshot (one fixed, one added via extendable port group).
        var portConfig = new ExcelCellUpdaterNodeFactory().createNodeCreationConfig().getPortConfig().get();
        portConfig.getExtendablePorts().get(ExcelCellUpdaterNodeFactory.SHEET_GRP_ID).addPort(BufferedDataTable.TYPE);
        return portConfig;
    }

    private static PortObjectSpec[] createInputPortSpecs() {
        // The node has one mandatory data table input port (plus optionally more via extendable port group).
        // We provide two table specs here to match the two sheet entries in the settings XML
        // (sheet_names array-size = 2, coordinate_column_names array-size = 2).
        return new PortObjectSpec[]{createDefaultTestTableSpec("id"), createDefaultTestTableSpec("address")};
    }

    private static DataTableSpec createDefaultTestTableSpec(final String idColumnName) {
        return new DataTableSpec( //
            new DataColumnSpecCreator(idColumnName, StringCell.TYPE).createSpec(), //
            new DataColumnSpecCreator("intValue", IntCell.TYPE).createSpec(), //
            new DataColumnSpecCreator("stringValue", StringCell.TYPE).createSpec());
    }

    @Test
    @SuppressWarnings("static-method")
    void testSheetNamesAreLoadedFromUnencryptedExcelFile() throws IOException, URISyntaxException {
        // given
        var settings = new ExcelCellUpdaterNodeParameters();
        settings.m_srcFileSelection =
            new FileSelection(new FSLocation(FSCategory.LOCAL, getTestFilePath(Fixtures.XLSX)));
        var context = NodeParametersUtil.createDefaultNodeSettingsContext(createInputPortSpecs(), createPortConfig());
        var simulator = new DialogUpdateSimulator(settings, context);

        // when
        var result = simulator.simulateValueChange("srcFileSelection");

        // then
        assertThat(
            result.getUiStateUpdateInArrayAt(List.of(List.of("sheetUpdates"), List.of("sheetName")), "possibleValues"))
                .isEqualTo(List.of(StringChoice.fromId("knime"), StringChoice.fromId("knime2")));
    }

    @Test
    @SuppressWarnings("static-method")
    void testSheetNamesAreLoadedFromPasswordProtectedExcelFile() throws IOException, URISyntaxException {
        // given
        var settings = new ExcelCellUpdaterNodeParameters();
        settings.m_srcFileSelection =
            new FileSelection(new FSLocation(FSCategory.LOCAL, getTestFilePath(Fixtures.XLSX_ENC)));
        settings.m_encryption = new ExcelEncryptionSettings(ExcelEncryptionSettings.AuthType.PWD, null,
            new Credentials("", Fixtures.TEST_PW));
        var context = NodeParametersUtil.createDefaultNodeSettingsContext(createInputPortSpecs(), createPortConfig());
        var simulator = new DialogUpdateSimulator(settings, context);

        // when
        var result = simulator.simulateValueChange("srcFileSelection");

        // then
        assertThat(
            result.getUiStateUpdateInArrayAt(List.of(List.of("sheetUpdates"), List.of("sheetName")), "possibleValues"))
                .isEqualTo(List.of(StringChoice.fromId("knime"), StringChoice.fromId("knime2")));
    }

    private static String getTestFilePath(final String path) throws IOException, URISyntaxException {
        var url = FileLocator.toFileURL(ExcelCellUpdaterNodeParametersTest.class.getResource(path).toURI().toURL());
        return FileUtil.getFileFromURL(url).toPath().toString();
    }
}
