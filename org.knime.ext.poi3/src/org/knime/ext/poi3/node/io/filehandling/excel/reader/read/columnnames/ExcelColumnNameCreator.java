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
 *   17.03.2021 (Lars Schweikardt, KNIME GmbH, Konstanz): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.columnnames;

import java.util.Set;

import org.knime.core.util.UniqueNameGenerator;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.AreaOfSheetToRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderTableSpec;

/**
 * A {@link ColumnNameCreator} to create column names based on the excel column names.
 *
 * @author Lars Schweikardt, KNIME GmbH, Konstanz
 */
final class ExcelColumnNameCreator implements ColumnNameCreator {

    private final TableReadConfig<ExcelTableReaderConfig> m_config;

    private final Set<Integer> m_hiddenColumns;

    private final String m_emptyPrefix;

    private final boolean m_useColHeaderIdx;

    private final UniqueNameGenerator m_uniqueNameGenerator;

    /**
     * Constructor.
     *
     * @param config the {@link TableReadConfig}
     * @param hiddenColumns a {@link Integer} {@link Set} of hidden columns
     * @param spec the {@link TypedReaderTableSpec}
     */
    public ExcelColumnNameCreator(final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns, final TypedReaderTableSpec<KNIMECellType> spec) {
        m_config = config;
        m_hiddenColumns = hiddenColumns;
        m_emptyPrefix = m_config.getReaderSpecificConfig().getEmptyColHeaderPrefix();
        m_useColHeaderIdx = m_config.useColumnHeaderIdx();
        m_uniqueNameGenerator = ExcelColNameUtils.createUniqueNameGenerator(spec);
    }

    @Override
    public String createColumnName(final int index) {
        final String columnName = (m_useColHeaderIdx ? m_emptyPrefix : "")
            + ExcelUtils.getExcelColumnName(getFilteredIdx(index, m_config, m_hiddenColumns));
        return m_uniqueNameGenerator.newName(columnName);
    }

    private static int getFilteredIdx(final int idx, final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns) {
        int filteredIdx;
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        final int startColIdx;
        // get the start index
        if (excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL) {
            startColIdx = getColIntervalStart(config);
            filteredIdx = idx + startColIdx;
        } else {
            startColIdx = 0;
            filteredIdx = idx;
        }
        final int rowIdColIdx = ExcelUtils.getRowIdColIdx(excelConfig.getRowIDCol());
        if (excelConfig.isSkipHiddenCols()) {
            // increment the index by the number of hidden columns that are before it and >= start
            final int i = filteredIdx;
            int count = (int)hiddenColumns.stream().filter(hIdx -> hIdx >= startColIdx && hIdx <= i).count();
            filteredIdx = filteredIdx + count;
            // increment as long as we have an index that is hidden
            while (hiddenColumns.contains(filteredIdx)) {
                filteredIdx++;
            }
            // increment if we had a row ID before the index and >= start (as this was filtered out and not counted)
            // make sure the row ID was not hidden as we have already covered that case above
            if (rowIdBeforeIdxAndWithinStart(config.useRowIDIdx(), filteredIdx, startColIdx, rowIdColIdx)
                && !hiddenColumns.contains(rowIdColIdx)) {
                filteredIdx++;
            }
        } else if (!excelConfig.isUseRawSettings()
            && rowIdBeforeIdxAndWithinStart(config.useRowIDIdx(), filteredIdx, startColIdx, rowIdColIdx)) {
            // increment if we had a row ID before the index and >= start (as this was filtered out and not counted)
            filteredIdx++;
        }
        return filteredIdx;
    }

    private static boolean rowIdBeforeIdxAndWithinStart(final boolean useRowIdIdx, final int filteredIdx,
        final int startColIdx, final int rowIdColIdx) {
        return useRowIdIdx && rowIdColIdx <= filteredIdx && rowIdColIdx >= startColIdx;
    }

    private static int getColIntervalStart(final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        if (excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.ENTIRE) {
            return 0;
        }
        return ExcelUtils.getFirstColumnIdx(excelConfig.getReadFromCol());
    }
}
