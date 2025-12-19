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

import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.HorizontalLayout;
import org.knime.node.parameters.layout.Inside;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;

/**
 * Parameters for skipping empty/hidden rows and columns in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Layout(SkipParameters.SkipLayout.class)
class SkipParameters implements NodeParameters {

    @Inside(ReaderLayout.DataArea.class)
    @After(RowIdParameters.RowIdLayout.class)
    interface SkipLayout {
        @HorizontalLayout
        interface SkipColumns {
        }

        @HorizontalLayout
        interface SkipRows {
        }
    }

    static final class SkipEmptyColumnsRef implements ParameterReference<Boolean> {
    }

    static final class SkipHiddenColumnsRef implements ParameterReference<Boolean> {
    }

    static final class SkipEmptyRowsRef implements ParameterReference<Boolean> {
    }

    static final class SkipHiddenRowsRef implements ParameterReference<Boolean> {
    }

    @Widget(title = "Skip empty columns",
        description = "If checked, empty columns of the sheet will be skipped and not displayed in the output. "
            + "Whether a column is considered empty depends on the <i>Column and Data Type Detection</i> settings: "
            + "If the cells of a column for all scanned rows were empty, the column is considered "
            + "empty, even if the sheet contains a non-empty cell after the scanned rows. "
            + "Removing the limit of scanned rows ensures empty and non-empty columns being detected correctly "
            + "but also increases the time required for scanning.")
    @Layout(SkipLayout.SkipColumns.class)
    @ValueReference(SkipEmptyColumnsRef.class)
    boolean m_skipEmptyColumns;

    @Widget(title = "Skip hidden columns",
        description = "If checked, hidden columns of the sheet will be skipped and not displayed in the output.")
    @Layout(SkipLayout.SkipColumns.class)
    @ValueReference(SkipHiddenColumnsRef.class)
    boolean m_skipHiddenColumns = true;

    @Widget(title = "Skip empty rows",
        description = "If checked, empty rows of the sheet will be skipped and not displayed in the output.")
    @Layout(SkipLayout.SkipRows.class)
    @ValueReference(SkipEmptyRowsRef.class)
    boolean m_skipEmptyRows = true;

    @Widget(title = "Skip hidden rows",
        description = "If checked, hidden rows of the sheet will be skipped and not displayed in the output.")
    @Layout(SkipLayout.SkipRows.class)
    @ValueReference(SkipHiddenRowsRef.class)
    boolean m_skipHiddenRows = true;

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var tableReadConfig = config.getTableReadConfig();

        excelConfig.setSkipHiddenCols(m_skipHiddenColumns);
        excelConfig.setSkipHiddenRows(m_skipHiddenRows);
        tableReadConfig.setSkipEmptyRows(m_skipEmptyRows);
        config.setSkipEmptyColumns(m_skipEmptyColumns);
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var tableReadConfig = config.getTableReadConfig();

        m_skipHiddenColumns = excelConfig.isSkipHiddenCols();
        m_skipHiddenRows = excelConfig.isSkipHiddenRows();
        m_skipEmptyRows = tableReadConfig.skipEmptyRows();
        m_skipEmptyColumns = config.skipEmptyColumns();
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // No specific validation needed for boolean flags
    }

}
