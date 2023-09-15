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
 *   14 Sep 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.CellConsumer;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.WorkbookMetadata;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;

/**
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class SSFParser implements XLParser {

    private Workbook m_workbook;
    private TableReadConfig<ExcelTableReaderConfig> m_config;
    private CompletableFuture<WorkbookMetadata> m_metadata;
    private CellConsumer m_output;
    private Path m_path;
    private Map<String, Boolean> m_sheetNames;
    private long m_numMaxRows;
    private boolean m_use1904windowing;

    private Sheet m_sheet;
    private ExcelCell m_rowId;
    private boolean m_skipHiddenRows;

    SSFParser(final Workbook workbook, final TableReadConfig<ExcelTableReaderConfig> config,
        final CompletableFuture<WorkbookMetadata> metadata, final CellConsumer output, final Path path) throws IOException {
        m_workbook = workbook;
        m_config = config;
        m_metadata = metadata;
        m_output = output;
        m_path = path;

        final FormulaEvaluator formulaEvaluator = m_config.getReaderSpecificConfig().isReevaluateFormulas()
                ? m_workbook.getCreationHelper().createFormulaEvaluator() : null;

        m_sheetNames = ExcelUtils.getSheetNames(m_workbook);
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        final var selectedSheet = excelConfig.getSheetSelection().getSelectedSheet(m_sheetNames, excelConfig, m_path);
        m_sheet = m_workbook.getSheet(selectedSheet);
        m_use1904windowing = use1904Windowing(m_workbook);
        m_numMaxRows = m_sheet.getLastRowNum() + 1L;

        m_skipHiddenRows = excelConfig.isSkipHiddenRows();

        // TODO metadata!
    }

    @Override
    public void close() throws IOException {
        m_workbook.close();
    }

    /**
     * Excel dates are encoded as a numeric offset of either 1900 or 1904.
     */
    private static boolean use1904Windowing(final Workbook workbook) {
        boolean date1904;
        if (workbook instanceof XSSFWorkbook xssfWorkbook) {
            date1904 = xssfWorkbook.isDate1904();
        } else if (workbook instanceof HSSFWorkbook hssfWorkbook) {
            date1904 = hssfWorkbook.getInternalWorkbook().isUsing1904DateWindowing();
        } else {
            // Probably unsupported
            date1904 = false;
        }
        return date1904;
    }

    @Override
    public void parse() {
        int lastNonEmptyRowIdx = -1;
        for (int i = 0; i <= m_sheet.getLastRowNum(); i++) {
            m_rowId = null;
            final Row row = m_sheet.getRow(i);

            if (row != null) {
                final boolean isHiddenRow = m_skipHiddenRows && row.getZeroHeight();
                // parse the row
                final List<ExcelCell> cells = parseRow(row);
                // if all cells of the row are null, the row is empty
                if (!ExcelUtils.isRowEmpty(cells) || m_rowId != null) {
                    outputEmptyRows(i - lastNonEmptyRowIdx - 1);
                    lastNonEmptyRowIdx = i;
                    // insert the row id at the beginning
                    insertRowIDAtBeginning(cells, m_rowId);
                    addToQueue(cells, isHiddenRow);
                }
            }
        }
        outputEmptyRows(m_lastRowIdx - lastNonEmptyRowIdx);
    }
}
