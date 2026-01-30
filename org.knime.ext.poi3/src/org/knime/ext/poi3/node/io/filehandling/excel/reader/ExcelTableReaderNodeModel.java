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
 *   Dec 7, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.util.function.Predicate;
import java.util.function.Supplier;

import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionPath;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.streamable.PartitionInfo;
import org.knime.core.node.streamable.StreamableOperator;
import org.knime.core.node.streamable.StreamableOperatorInternals;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.node.table.reader.BackwardsCompatibleCommonTableReaderNodeModel;
import org.knime.filehandling.core.node.table.reader.CommonTableReaderNodeFactory;
import org.knime.filehandling.core.node.table.reader.MultiTableReader;
import org.knime.filehandling.core.node.table.reader.config.StorableMultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.paths.SourceSettings;

/**
 * Custom node model for the {@link ExcelTableReaderNodeFactory} to set the credentials provider
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
final class ExcelTableReaderNodeModel extends
    BackwardsCompatibleCommonTableReaderNodeModel<FSPath, MultiFileSelectionPath, ExcelTableReaderConfig, KNIMECellType, ExcelMultiTableReadConfig> {

    ExcelTableReaderNodeModel(final Supplier<ExcelMultiTableReadConfig> config,
        final MultiFileSelectionPath pathSettingsModel,
        final MultiTableReader<FSPath, ExcelTableReaderConfig, KNIMECellType> tableReader,
        final CommonTableReaderNodeFactory.ConfigAndSourceSerializer<FSPath, MultiFileSelectionPath, ExcelTableReaderConfig, KNIMECellType, ExcelMultiTableReadConfig> serializer,
        final SourceSettings<FSPath> legacySourceSettings, final Predicate<NodeSettingsRO> isLegacySettingsPredicate) {
        super(config, pathSettingsModel, tableReader, serializer, legacySourceSettings, isLegacySettingsPredicate);
    }

    ExcelTableReaderNodeModel(final Supplier<ExcelMultiTableReadConfig> config,
        final MultiFileSelectionPath pathSettingsModel,
        final MultiTableReader<FSPath, ExcelTableReaderConfig, KNIMECellType> tableReader,
        final CommonTableReaderNodeFactory.ConfigAndSourceSerializer<FSPath, MultiFileSelectionPath, ExcelTableReaderConfig, KNIMECellType, ExcelMultiTableReadConfig> serializer,
        final PortsConfiguration portsConfig, final SourceSettings<FSPath> legacySourceSettings,
        final Predicate<NodeSettingsRO> isLegacySettingsPredicate) {
        super(config, pathSettingsModel, tableReader, serializer, portsConfig, legacySourceSettings,
            isLegacySettingsPredicate);
    }

    @Override
    protected PortObject[] execute(final PortObject[] inObjects, final ExecutionContext exec) throws Exception {
        setCredentialsProvider();
        return super.execute(inObjects, exec);
    }

    @Override
    public PortObjectSpec[] computeFinalOutputSpecs(final StreamableOperatorInternals internals,
        final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        setCredentialsProvider();
        return super.computeFinalOutputSpecs(internals, inSpecs);
    }

    @Override
    public StreamableOperator createStreamableOperator(final PartitionInfo partitionInfo,
        final PortObjectSpec[] inSpecs) throws InvalidSettingsException {
        setCredentialsProvider();
        return super.createStreamableOperator(partitionInfo, inSpecs);
    }

    private void setCredentialsProvider() {
        final StorableMultiTableReadConfig<ExcelTableReaderConfig, KNIMECellType> config = getConfig();
        final ExcelTableReaderConfig excelConfig = config.getTableReadConfig().getReaderSpecificConfig();
        excelConfig.setCredentialsProvider(getCredentialsProvider());
    }
}
