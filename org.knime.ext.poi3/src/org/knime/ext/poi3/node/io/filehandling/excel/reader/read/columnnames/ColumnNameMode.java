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
 *   Mar 10, 2020 (Lars Schweikardt, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.columnnames;

import java.util.Set;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderTableSpec;

/**
 * Enumeration for the column name generation settings.
 *
 * @author Lars Schweikardt, KNIME GmbH, Konstanz, Germany
 */
public enum ColumnNameMode {

        /** Creates column headers with a number starting from 0 */
        COL_INDEX("Use column index e.g. Col0, Col1, Col2") {
            @Override
            public ColumnNameCreator getColumnNameCreator(final TableReadConfig<ExcelTableReaderConfig> config,
                final Set<Integer> hiddenColumns, final TypedReaderTableSpec<KNIMECellType> spec) {
                return new IndexColumnNameCreator(config, spec);
            }
        },
        /** Creates column headers with the corresponding excel column name */
        EXCEL_COL_NAME("Use Excel column name e.g. A, B, C") {
            @Override
            public ColumnNameCreator getColumnNameCreator(final TableReadConfig<ExcelTableReaderConfig> config,
                final Set<Integer> hiddenColumns, final TypedReaderTableSpec<KNIMECellType> spec) {
                return new ExcelColumnNameCreator(config, hiddenColumns, spec);
            }
        };

    private static final NodeLogger LOGGER =  NodeLogger.getLogger(ColumnNameMode.class);

    private final String m_text;

    private ColumnNameMode(final String text) {
        m_text = text;
    }

    /**
     * Returns the description of a mode.
     *
     * @return the description of a mode
     */
    public String getText() {
        return m_text;
    }

    /**
     * Abstract method which returns a instance of a {@link ColumnNameCreator}.
     *
     * @param config the {@link TableReadConfig}
     * @param hiddenColumns a {@link Integer} {@link Set} of hidden columns
     * @return a concrete instance of a {@link ColumnNameCreator}
     */
    abstract ColumnNameCreator getColumnNameCreator(final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns, final TypedReaderTableSpec<KNIMECellType> spec);

    public static ColumnNameMode loadValueInModel(final String s) throws InvalidSettingsException {
        try {
            return valueOf(s);
        } catch (IllegalArgumentException e) {
            LOGGER.debug(String.format("%s is no valid column name mode", s), e);
            throw new InvalidSettingsException(
                "No valid column name mode '" + s + "' available. See node description.");
        }
    }

}
