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
 *   Nov 16, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.sheets;

import java.io.IOException;
import java.util.Optional;

import org.apache.xmlbeans.XmlException;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ConfigurableNodeFactory;
import org.knime.core.node.NodeDescription;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeFactory;
import org.knime.core.node.NodeView;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.webui.node.dialog.NodeDialog;
import org.knime.core.webui.node.dialog.NodeDialogFactory;
import org.knime.core.webui.node.dialog.NodeDialogManager;
import org.knime.core.webui.node.dialog.SettingsType;
import org.knime.core.webui.node.dialog.defaultdialog.DefaultNodeDialog;
import org.knime.core.webui.node.impl.WebUINodeConfiguration;
import org.knime.core.webui.node.impl.WebUINodeFactory;
import org.knime.filehandling.core.defaultnodesettings.EnumConfig;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.port.FileSystemPortObject;
import org.xml.sax.SAXException;

/**
 * The {@link NodeFactory} of the 'Read Excel Sheet Names' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
@SuppressWarnings({"restriction", "deprecation"})
public final class ExcelSheetReaderNodeFactory extends ConfigurableNodeFactory<ExcelSheetReaderNodeModel>
    implements NodeDialogFactory {

    /** The file system ports group id. */
    static final String FS_CONNECT_GRP_ID = "File System Connection";

    /** The supported file extensions. */
    private static final String[] FILE_EXTENSIONS = new String[]{".xlsx", ".xlsm", ".xlsb", ".xls"};

    @Override
    protected Optional<PortsConfigurationBuilder> createPortsConfigBuilder() {
        final PortsConfigurationBuilder b = new PortsConfigurationBuilder();
        b.addOptionalInputPortGroup(FS_CONNECT_GRP_ID, FileSystemPortObject.TYPE);
        b.addFixedOutputPortGroup("Output table", BufferedDataTable.TYPE);
        return Optional.of(b);
    }

    @Override
    protected ExcelSheetReaderNodeModel createNodeModel(final NodeCreationConfiguration creationConfig) {
        final PortsConfiguration portsConfig = creationConfig.getPortConfig().orElseThrow(IllegalStateException::new);
        return new ExcelSheetReaderNodeModel(portsConfig, createFileChooser(portsConfig));
    }

    private static final WebUINodeConfiguration CONFIG = WebUINodeConfiguration.builder()//
        .name("Read Excel Sheet Names")//
        .icon("excel_sheet_reader.png")//
        .shortDescription("Reads the sheet names contained in an Excel file.")//
        .fullDescription(
            """
                    <p>
                    This node reads an Excel file and provides the contained sheet names at its output port.<br />
                    The performance of the reader node is limited (due to the underlying library
                    of the Apache POI project). Reading large files can take a long time.
                    </p>

                    <p>
                    <i>This node can access a variety of different</i>
                    <a href="https://docs.knime.com/2021-06/analytics_platform_file_handling_guide/index.html#analytics-platform-file-systems"><i>file systems.</i></a>
                    <i>More information about file handling in KNIME can be found in the official</i>
                    <a href="https://docs.knime.com/latest/analytics_platform_file_handling_guide/index.html"><i>File Handling Guide.</i></a>
                    </p>
                    """)//
        .modelSettingsClass(ExcelSheetReaderNodeParameters.class)//
        .addInputPort(FS_CONNECT_GRP_ID, FileSystemPortObject.TYPE,
            "The file system connection providing access to the selected file(s).", true)//
        .addOutputPort("Output table", BufferedDataTable.TYPE,
            "The sheet names contained in the workbook and the path to the workbook.")//
        .nodeType(NodeType.Source)//
        .keywords(new String[]{"excel", "sheet", "read"})//
        .sinceVersion(4, 5, 0)// best-effort (node originally introduced earlier)
        .build();

    @Override
    protected NodeDescription createNodeDescription() throws SAXException, IOException, XmlException {
        return WebUINodeFactory.createNodeDescription(CONFIG);
    }

    @Override
    protected NodeDialogPane createNodeDialogPane(final NodeCreationConfiguration creationConfig) {
        // keep legacy flow variable support wrapper
        return NodeDialogManager.createLegacyFlowVariableNodeDialog(createNodeDialog());
    }

    @Override
    protected int getNrNodeViews() {
        return 0;
    }

    @Override
    public NodeView<ExcelSheetReaderNodeModel> createNodeView(final int viewIndex,
        final ExcelSheetReaderNodeModel nodeModel) {
        return null;
    }

    @Override
    protected boolean hasDialog() {
        return true;
    }

    // --- Modern UI additions -------------------------------------------------

    @Override
    public NodeDialog createNodeDialog() {
        return new DefaultNodeDialog(SettingsType.MODEL, ExcelSheetReaderNodeParameters.class);
    }

    private static SettingsModelReaderFileChooser createFileChooser(final PortsConfiguration portsConfig) {
        return new SettingsModelReaderFileChooser("file_selection", portsConfig, FS_CONNECT_GRP_ID,
            EnumConfig.create(FilterMode.FILE, FilterMode.FILES_IN_FOLDERS), FILE_EXTENSIONS);
    }

}
