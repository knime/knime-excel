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

import static org.knime.node.impl.description.PortDescription.dynamicPort;
import static org.knime.node.impl.description.PortDescription.fixedPort;

import java.util.List;

import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionPath;
import org.knime.base.node.io.filehandling.webui.reader2.NodeParametersConfigAndSourceSerializer;
import org.knime.base.node.io.filehandling.webui.reader2.WebUITableReaderNodeFactory;
import org.knime.core.node.NodeDescription;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.util.Version;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.node.table.reader.GenericTableReader;
import org.knime.filehandling.core.node.table.reader.ReadAdapterFactory;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeHierarchy;
import org.knime.node.impl.description.DefaultNodeDescriptionUtil;

/**
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
public class ExcelTableReaderNodeFactory extends WebUITableReaderNodeFactory<ExcelTableReaderNodeParameters, FSPath, //
        MultiFileSelectionPath, ExcelTableReaderConfig, KNIMECellType, ExcelCell, ExcelMultiTableReadConfig> {

    @SuppressWarnings("javadoc")
    public ExcelTableReaderNodeFactory() {
        super(ExcelTableReaderNodeParameters.class);
    }

    // TODO (#5): Complete the node description
    @Override
    protected NodeDescription createNodeDescription() {
        return DefaultNodeDescriptionUtil.createNodeDescription( //
            "TODO (#5): Name", //
            "TODO (#5): Icon", //
            List.of(dynamicPort(FS_CONNECT_GRP_ID, "File System Connection", "The file system connection.")), //
            List.of(fixedPort("TODO (#5): Output name", "TODO (#5): Output description")), //
            "TODO (#5): short description", //
            "TODO (#5): Full description", //
            List.of(), //
            ExcelTableReaderNodeParameters.class, //
            null, //
            NodeType.Source, //
            List.of("Input", "Read"), // TODO (#5): set additional keywords
            new Version(5, 9, 0) // TODO (#5): Set the version when the node is first introduced
        );
    }

    // TODO (#5): Return your ReadAdapterFactory instance
    @Override
    protected ReadAdapterFactory<KNIMECellType, ExcelCell> getReadAdapterFactory() {
        return null;
    }

    @Override
    protected GenericTableReader<FSPath, ExcelTableReaderConfig, KNIMECellType, ExcelCell> createReader() {
        return new ExcelTableReader();
    }

    // TODO (#5): Implement if your V type requires special handling
    @Override
    protected String extractRowKey(final ExcelCell value) {
        return value.getStringValue();
    }

    // TODO (#5): Return your ReadAdapterFactory's TYPE_HIERARCHY
    @Override
    protected TypeHierarchy<KNIMECellType, KNIMECellType> getTypeHierarchy() {
        return null;
    }

    @Override
    protected ExcelTableReaderConfigAndSourceSerializer createSerializer() {
        return new ExcelTableReaderConfigAndSourceSerializer();
    }

    private final class ExcelTableReaderConfigAndSourceSerializer
        extends NodeParametersConfigAndSourceSerializer<ExcelTableReaderNodeParameters, FSPath, MultiFileSelectionPath, //
                ExcelTableReaderConfig, KNIMECellType, ExcelMultiTableReadConfig> {
        protected ExcelTableReaderConfigAndSourceSerializer() {
            super(ExcelTableReaderNodeParameters.class);
        }

        @Override
        protected void saveToSourceAndConfig(final ExcelTableReaderNodeParameters params,
            final MultiFileSelectionPath sourceSettings, final ExcelMultiTableReadConfig config) {
            params.saveToSource(sourceSettings);
            params.saveToConfig(config);
        }
    }

    @Override
    protected MultiFileSelectionPath createPathSettings(final NodeCreationConfiguration nodeCreationConfig) {
        final var source = new MultiFileSelectionPath();
        final var defaultParams = new ExcelTableReaderNodeParameters(nodeCreationConfig);
        defaultParams.saveToSource(source);
        return source;
    }

    @Override
    protected ExcelMultiTableReadConfig createConfig(final NodeCreationConfiguration nodeCreationConfig) {
        final var cfg = new ExcelMultiTableReadConfig();
        final var defaultParams = new ExcelTableReaderNodeParameters(nodeCreationConfig);
        defaultParams.saveToConfig(cfg);
        return cfg;
    }
}
