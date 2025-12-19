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
 *   Dec 9, 2020 (Adrian Nembach, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.node.parameters.migration.ConfigMigration;
import org.knime.node.parameters.migration.NodeParametersMigration;

import java.util.List;

/**
 * Migrate from legacy settings to the new {@link ExcelTableReaderNodeParameters}.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
class ExcelTableReaderMigration implements NodeParametersMigration<ExcelTableReaderNodeParameters> {

    private static final String SETTINGS_KEY = "settings";

    private static final String ADVANCED_KEY = "advanced_settings";

    private static final String TABLE_TRANSFORMATION_KEY = "table_spec_config_Internals";

    private static final String ENCRYPTION_KEY = "encryption";

    private static ExcelTableReaderNodeParameters load(final NodeSettingsRO settings) throws InvalidSettingsException {
        var newSettings = new ExcelTableReaderNodeParameters();
        var config = new ExcelMultiTableReadConfig();
        ExcelMultiTableReadConfigSerializer.INSTANCE.loadInModel(config, settings);

        newSettings.m_excelTableReaderParameters.m_multiFileSelectionParams.loadFromLegacySettings(settings);
        newSettings.m_excelTableReaderParameters.m_ifSchemaChangesParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_multiFileReaderParams.loadFromConfigAfter4_4(config);
        newSettings.m_excelTableReaderParameters.m_selectSheetParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_readAreaParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_columnNameParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_rowIdParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_skipParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_schemaDetectionParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_valuesParams.loadFromConfig(config);
        newSettings.m_excelTableReaderParameters.m_encryptionParams.loadFromConfig(config);
        newSettings.m_transformationParameters.loadFromLegacySettings(settings);

        return newSettings;
    }

    @Override
    public List<ConfigMigration<ExcelTableReaderNodeParameters>> getConfigMigrations() {
        return List.of(//
            ConfigMigration.builder(ExcelTableReaderMigration::load) //
                .withMatcher(s -> s.containsKey(SETTINGS_KEY))
                .withDeprecatedConfigPath(SETTINGS_KEY)//
                .withDeprecatedConfigPath(ADVANCED_KEY)//
                .withDeprecatedConfigPath(TABLE_TRANSFORMATION_KEY) //
                .withDeprecatedConfigPath(ENCRYPTION_KEY) //
                .build());
    }
}
