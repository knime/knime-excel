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
 *   Nov 6, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import static org.knime.node.impl.description.PortDescription.dynamicPort;
import static org.knime.node.impl.description.PortDescription.fixedPort;

import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ConfigurableNodeFactory;
import org.knime.core.node.NodeDescription;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeFactory;
import org.knime.core.node.NodeView;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.port.PortType;
import org.knime.core.webui.node.dialog.NodeDialog;
import org.knime.core.webui.node.dialog.NodeDialogFactory;
import org.knime.core.webui.node.dialog.NodeDialogManager;
import org.knime.core.webui.node.dialog.SettingsType;
import org.knime.core.webui.node.dialog.defaultdialog.DefaultKaiNodeInterface;
import org.knime.core.webui.node.dialog.defaultdialog.DefaultNodeDialog;
import org.knime.core.webui.node.dialog.kai.KaiNodeInterface;
import org.knime.core.webui.node.dialog.kai.KaiNodeInterfaceFactory;
import org.knime.filehandling.core.port.FileSystemPortObject;
import org.knime.node.impl.description.DefaultNodeDescriptionUtil;
import org.knime.node.impl.description.PortDescription;

/**
 * {@link NodeFactory} creating the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline v1.2
 */
@SuppressWarnings("restriction")
public final class ExcelTableWriterNodeFactory extends ConfigurableNodeFactory<ExcelTableWriterNodeModel>
    implements NodeDialogFactory, KaiNodeInterfaceFactory {

    /** The file system ports group id. */
    static final String FS_CONNECT_GRP_ID = "File System Connection";

    /** The sheet ports group id. */
    static final String SHEET_GRP_ID = "Sheet Input Ports";

    @Override
    protected Optional<PortsConfigurationBuilder> createPortsConfigBuilder() {
        final var b = new PortsConfigurationBuilder();
        b.addOptionalInputPortGroup(FS_CONNECT_GRP_ID, FileSystemPortObject.TYPE);
        b.addExtendableInputPortGroup(SHEET_GRP_ID, new PortType[]{BufferedDataTable.TYPE}, BufferedDataTable.TYPE);
        return Optional.of(b);
    }

    @Override
    protected ExcelTableWriterNodeModel createNodeModel(final NodeCreationConfiguration creationConfig) {
        return new ExcelTableWriterNodeModel(creationConfig.getPortConfig().orElseThrow(IllegalStateException::new));
    }

    @Override
    protected int getNrNodeViews() {
        return 0;
    }

    @Override
    @SuppressWarnings("removal")
    public NodeView<ExcelTableWriterNodeModel> createNodeView(final int viewIndex,
        final ExcelTableWriterNodeModel nodeModel) {
        return null;
    }

    @Override
    @SuppressWarnings("removal")
    protected boolean hasDialog() {
        return true;
    }

    private static final String NODE_NAME = "Excel Writer";

    private static final String NODE_ICON = "./excel_writer.png";

    private static final String SHORT_DESCRIPTION = """
            Writes data tables to an Excel file.
            """;

    @SuppressWarnings("java:S103")
    private static final String FULL_DESCRIPTION =
        """
                <p> This node writes the input data table into a spreadsheet of an Excel file, which can then be read
                    with other applications such as Microsoft Excel. The node can create completely new files or append data
                    to an existing Excel file. When appending, the input data can be appended as a new spreadsheet or after
                    the last row of an existing spreadsheet. By adding multiple data table input ports, the data can be
                    written/appended to multiple spreadsheets within the same file. </p> <p>The node supports two formats
                    chosen by file extension: <ul> <li> <tt>.xls</tt> format: This is the file format which was used by
                    default up until Excel 2003. The maximum number of columns and rows held by a spreadsheet of this format
                    is 256 and 65536 respectively. </li> <li> <tt>.xlsx</tt> format: The Office Open XML format is the file
                    format used by default from Excel 2007 onwards. The maximum number of columns and rows held by a
                    spreadsheet of this format is 16384 and 1048576 respectively. </li> </ul> </p> <p> If the data does not
                    fit into a single sheet, it will be split into multiple chunks that are written to newly chosen sheet
                    names sequentially. The new sheet names are derived from the originally selected sheet name by appending
                    " (i)" to it, where i=1,...,n. <br /> When appending to a file, a sheet may already exist. In this case
                    the node will (according to its settings) either replace the sheet, fail or append rows after the last
                    row in that sheet. This can be used to append data to an Excel file without having to create a new sheet
                    name once the original sheet is full by just selecting the original sheet name. The data will be
                    appendend to the last sheet in the name sequence. The original sheet name does not have to be changed.
                    </p> <p>This node does not support writing files in the '.xlsm' format, yet appending is supported.</p>
                    <p> <i>This node can access a variety of different</i> <a
                    href="https://docs.knime.com/ap/latest/analytics_platform_file_handling_guide/index.html#knime-analytics-platform-and-file-systems"><i>file
                    systems.</i></a> <i>More information about file handling in KNIME can be found in the official</i> <a
                    href="https://docs.knime.com/latest/analytics_platform_file_handling_guide/index.html"><i>File Handling
                    Guide.</i></a> </p>
                """;

    private static final List<PortDescription> INPUT_PORTS =
        List.of(dynamicPort(FS_CONNECT_GRP_ID, "File system connection", """
                The file system connection.
                """), fixedPort("Input table", """
                The data table to write.
                """), dynamicPort(SHEET_GRP_ID, "Additional input tables", """
                Additional data table to write.
                """));

    private static final List<PortDescription> OUTPUT_PORTS = List.of();

    private static final List<String> KEYWORDS = List.of( //
        "Spreadsheet", //
        "XLS Writer", //
        "Microsoft", //
        "write excel file" //
    );

    @Override
    @SuppressWarnings("removal")
    public NodeDialogPane createNodeDialogPane(final NodeCreationConfiguration creationConfig) {
        return NodeDialogManager.createLegacyFlowVariableNodeDialog(createNodeDialog());
    }

    @Override
    public NodeDialog createNodeDialog() {
        return new DefaultNodeDialog(SettingsType.MODEL, ExcelTableWriterNodeParameters.class);
    }

    @Override
    public NodeDescription createNodeDescription() {
        return DefaultNodeDescriptionUtil.createNodeDescription(NODE_NAME, NODE_ICON, INPUT_PORTS, OUTPUT_PORTS,
            SHORT_DESCRIPTION, FULL_DESCRIPTION, List.of(), ExcelTableWriterNodeParameters.class, null, NodeType.Sink,
            KEYWORDS, null);
    }

    @Override
    public KaiNodeInterface createKaiNodeInterface() {
        return new DefaultKaiNodeInterface(Map.of(SettingsType.MODEL, ExcelTableWriterNodeParameters.class));
    }

}
