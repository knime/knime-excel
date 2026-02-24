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
 *  Aug 9, 2021 (Moditha Hewasinghage, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

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
 * {@link NodeFactory} creating the 'Excel Cell Updater' node.
 *
 * @author Moditha Hewasinghage, KNIME GmbH, Berlin, Germany
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 * @author AI Migration Pipeline v1.2
 */
@SuppressWarnings("restriction")
public final class ExcelCellUpdaterNodeFactory extends ConfigurableNodeFactory<ExcelCellUpdaterNodeModel>
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
    protected ExcelCellUpdaterNodeModel createNodeModel(final NodeCreationConfiguration creationConfig) {
        return new ExcelCellUpdaterNodeModel(creationConfig.getPortConfig().orElseThrow(IllegalStateException::new));
    }

    @Override
    protected int getNrNodeViews() {
        return 0;
    }

    @SuppressWarnings("removal")
    @Override
    public NodeView<ExcelCellUpdaterNodeModel> createNodeView(final int viewIndex,
        final ExcelCellUpdaterNodeModel nodeModel) {
        return null;
    }

    @SuppressWarnings("removal")
    @Override
    protected boolean hasDialog() {
        return true;
    }

    private static final String NODE_NAME = "Excel Cell Updater";

    private static final String NODE_ICON = "./excel_cell_updater.png";

    private static final String SHORT_DESCRIPTION = """
            Updates single cells in an Excel file
            """;

    private static final String FULL_DESCRIPTION =
        """
                <p> This node updates cells in an existing Excel spreadsheet. The cell addresses and their new content
                are supplied by an input data table. </p> <p> Adding additional table input ports allows to update
                multiple spreadsheets in the same file. </p> <p> The node supports two file formats chosen by file
                extension: <ul> <li> .xls format: This is the file format which was used by default up until Excel 2003.
                The maximum number of columns held by a spreadsheet of this format is 256 (address IV) and the maximum
                number of rows is 65536. </li> <li> .xlsx format: The Office Open XML format is the file format used by
                default from Excel 2007 onwards. The maximum number of columns held by a spreadsheet of this format is
                16384 (address XFD) and the maximum number of rows is 1048576. </li> </ul> </p> <p> Each input table
                must have one column which contains the addresses of the sheet cells which should be updated. This
                column has to have a string-compatible type. Two types of addresses are supported and can be used
                interchangeably: <ul> <li>Excel cell addresses, e.g. "A5", "E96" or "OZ23914"</li> <li> Number addresses
                separated by a colon ":" with the first number being the number of the column and the second the one of
                the row. Both start counting with 1. You can use the <tt>COLUMN()</tt> function in Excel or enable
                Excel's R1C1 reference style to get the column number. Examples: "1:5", "5:96" or "413:23914" </li>
                </ul> </p> <p> The remaining columns contain the replacement values for the specified cells. A
                replacement value should appear in at most one column per row and the remaining cells should be left
                empty (i.e. only contain missing values). The column type should be the same as the (desired) column
                type in the updated sheet. If all cells in a row but the cell address contain missing values, the
                replacement value is a blank cell or the string specified in the “Replace missing values by” field. </p>
                <p> The formatting of existing cells, rows and columns in the Excel sheet will be preserved. </p> <p> In
                the following table <b>?</b> stands for a missing value. This example table would write the string "Ok"
                to Excel cell A5 and the number 50 to cell E96 (in number address style). The cell OZ23914 will be
                cleared if no alternate missing value is defined. </p> <table> <tr> <th>Address</th> <th>String</th>
                <th>Integer</th> </tr> <tr> <td>A5</td> <td>Ok</td> <td><b>?</b></td> </tr> <tr> <td>5:96</td>
                <td><b>?</b></td> <td>50</td> </tr> <tr> <td>OZ23914</td> <td><b>?</b></td> <td><b>?</b></td> </tr>
                </table> <p> <i>This node can access a variety of different</i> <a
                href="https://docs.knime.com/2021-06/analytics_platform_file_handling_guide/index.html#analytics-platform-file-systems"><i>file
                systems.</i></a> <i>More information about file handling in KNIME can be found in the official</i> <a
                href="https://docs.knime.com/latest/analytics_platform_file_handling_guide/index.html"><i>File Handling
                Guide.</i></a> </p>
                """;

    private static final List<PortDescription> INPUT_PORTS =
        List.of(dynamicPort(FS_CONNECT_GRP_ID, "File system connection", """
                The file system connection.
                """), fixedPort("Input table", """
                The data table which contains the update information.
                """), dynamicPort(SHEET_GRP_ID, "Additional input tables", """
                Additional data tables which contain update data.
                """));

    private static final List<PortDescription> OUTPUT_PORTS = List.of();

    @Override
    @SuppressWarnings("removal")
    public NodeDialogPane createNodeDialogPane(final NodeCreationConfiguration creationConfig) {
        return NodeDialogManager.createLegacyFlowVariableNodeDialog(createNodeDialog());
    }

    @Override
    public NodeDialog createNodeDialog() {
        return new DefaultNodeDialog(SettingsType.MODEL, ExcelCellUpdaterNodeParameters.class);
    }

    @Override
    public NodeDescription createNodeDescription() {
        return DefaultNodeDescriptionUtil.createNodeDescription( //
            NODE_NAME, //
            NODE_ICON, //
            INPUT_PORTS, //
            OUTPUT_PORTS, //
            SHORT_DESCRIPTION, //
            FULL_DESCRIPTION, //
            List.of(), //
            ExcelCellUpdaterNodeParameters.class, //
            null, //
            NodeType.Sink, //
            List.of(), //
            null //
        );
    }

    @Override
    public KaiNodeInterface createKaiNodeInterface() {
        return new DefaultKaiNodeInterface(Map.of(SettingsType.MODEL, ExcelCellUpdaterNodeParameters.class));
    }

}
