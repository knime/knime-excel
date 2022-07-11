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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.util.Optional;

import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.context.url.URLConfiguration;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelReadAdapterFactory;
import org.knime.filehandling.core.connections.FSLocationUtil;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.defaultnodesettings.EnumConfig;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.node.table.reader.AbstractTableReaderNodeFactory;
import org.knime.filehandling.core.node.table.reader.HierarchyAwareProductionPathProvider;
import org.knime.filehandling.core.node.table.reader.MultiTableReadFactory;
import org.knime.filehandling.core.node.table.reader.MultiTableReader;
import org.knime.filehandling.core.node.table.reader.ProductionPathProvider;
import org.knime.filehandling.core.node.table.reader.ReadAdapterFactory;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.StorableMultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.paths.PathSettings;
import org.knime.filehandling.core.node.table.reader.preview.dialog.AbstractTableReaderNodeDialog;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TreeTypeHierarchy;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeHierarchy;

/**
 * Node factory of Excel reader node.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelTableReaderNodeFactory
    extends AbstractTableReaderNodeFactory<ExcelTableReaderConfig, KNIMECellType, ExcelCell> {

    private static final String[] FILE_SUFFIXES = new String[]{".xlsx", ".xlsm", ".xlsb", ".xls"};

    private static final TreeTypeHierarchy<KNIMECellType, KNIMECellType> TYPE_HIERARCHY =
        ExcelTableReader.TYPE_HIERARCHY.createTypeFocusedHierarchy();

    @Override
    public ExcelTableReaderNodeModel createNodeModel(final NodeCreationConfiguration creationConfig) {
        final StorableMultiTableReadConfig<ExcelTableReaderConfig, KNIMECellType> config = createConfig(creationConfig);
        final PathSettings pathSettings = createPathSettings(creationConfig);
        final MultiTableReader<FSPath, ExcelTableReaderConfig, KNIMECellType> reader = createMultiTableReader();
        final Optional<? extends PortsConfiguration> portConfig = creationConfig.getPortConfig();
        if (portConfig.isPresent()) {
            return new ExcelTableReaderNodeModel(config, pathSettings, reader, portConfig.get());
        } else {
            return new ExcelTableReaderNodeModel(config, pathSettings, reader);
        }
    }

    @Override
    protected ProductionPathProvider<KNIMECellType> createProductionPathProvider() {
        final ExcelReadAdapterFactory raf = ExcelReadAdapterFactory.INSTANCE;
        return new HierarchyAwareProductionPathProvider<>(raf.getProducerRegistry(), TYPE_HIERARCHY, raf::getDefaultType,
            (t, p) -> true);
    }

    @Override
    protected AbstractTableReaderNodeDialog<FSPath, ExcelTableReaderConfig, KNIMECellType> createNodeDialogPane(
        final NodeCreationConfiguration creationConfig,
        final MultiTableReadFactory<FSPath, ExcelTableReaderConfig, KNIMECellType> readFactory,
        final ProductionPathProvider<KNIMECellType> defaultProductionPathFn) {
        // we overwrite #createNodeDialogPane(NodeCreationConfiguration) directly to be able to pass a reference of the
        // ExcelTableReader to the dialog; this reference will be used to fetch sheet name information during spec
        // guessing
        return null;
    }

    @Override
    protected NodeDialogPane createNodeDialogPane(final NodeCreationConfiguration creationConfig) {
        final ExcelTableReader tableReader = createReader();
        final MultiTableReadFactory<FSPath, ExcelTableReaderConfig, KNIMECellType> readFactory =
            createMultiTableReadFactory(tableReader);
        final ProductionPathProvider<KNIMECellType> productionPathProvider = createProductionPathProvider();
        return new ExcelTableReaderNodeDialog(createPathSettings(creationConfig), createConfig(creationConfig),
            tableReader, readFactory, productionPathProvider);
    }

    @Override
    protected SettingsModelReaderFileChooser createPathSettings(final NodeCreationConfiguration nodeCreationConfig) {
        final SettingsModelReaderFileChooser settingsModel = new SettingsModelReaderFileChooser("file_selection",
            nodeCreationConfig.getPortConfig().orElseThrow(IllegalStateException::new), FS_CONNECT_GRP_ID,
            EnumConfig.create(FilterMode.FILE, FilterMode.FILES_IN_FOLDERS), FILE_SUFFIXES);
        final Optional<? extends URLConfiguration> urlConfig = nodeCreationConfig.getURLConfig();
        if (urlConfig.isPresent()) {
            settingsModel
                .setLocation(FSLocationUtil.createFromURL(urlConfig.get().getUrl().toString()));
        }
        return settingsModel;
    }

    @Override
    protected ReadAdapterFactory<KNIMECellType, ExcelCell> getReadAdapterFactory() {
        return ExcelReadAdapterFactory.INSTANCE;
    }

    @Override
    protected ExcelTableReader createReader() {
        return new ExcelTableReader();
    }

    @Override
    protected String extractRowKey(final ExcelCell value) {
        return value.getStringValue();
    }

    @Override
    protected TypeHierarchy<KNIMECellType, KNIMECellType> getTypeHierarchy() {
        return ExcelTableReader.TYPE_HIERARCHY.createTypeFocusedHierarchy();
    }

    @Override
    protected ExcelMultiTableReadConfig createConfig(final NodeCreationConfiguration nodeCreationConfig) {
        ExcelMultiTableReadConfig config = new ExcelMultiTableReadConfig();
        final DefaultTableReadConfig<ExcelTableReaderConfig> tc = config.getTableReadConfig();
        tc.setColumnHeaderIdx(0);
        tc.setSkipEmptyRows(true);
        tc.setAllowShortRows(true);
        tc.setDecorateRead(false);
        return config;
    }

    @Override
    protected boolean hasDialog() {
        return true;
    }

}
