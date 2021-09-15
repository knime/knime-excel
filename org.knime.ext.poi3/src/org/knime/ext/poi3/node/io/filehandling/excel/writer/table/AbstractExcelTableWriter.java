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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.table;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.streamable.RowInput;
import org.knime.core.util.Pair;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.sheet.ExcelSheetWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelProgressMonitor;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetNameExistsHandling;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetUtils;

/**
 * An abstract {@link ExcelTableWriter}, that allows to write an {@link RowInput} to a sheet. If the number of rows to
 * be written exceeds the sheet's limitations additional sheets are automatically created.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
abstract class AbstractExcelTableWriter implements ExcelTableWriter {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(AbstractExcelTableWriter.class);

    private final ExcelTableConfig m_cfg;

    private final ExcelCellWriterFactory m_cellWriterFactory;

    private final int m_maxNumRowsPerSheet;

    AbstractExcelTableWriter(final ExcelTableConfig cfg, final ExcelCellWriterFactory cellWriterFactory,
        final int maxNumRowsPerSheet) {
        m_cfg = cfg;
        m_cellWriterFactory = cellWriterFactory;
        m_maxNumRowsPerSheet = maxNumRowsPerSheet;
    }

    @Override
    public final void writeTable(final Workbook workbook, final String sheetName, final RowInput rowInput,
        final ExcelProgressMonitor monitor)
        throws IOException, InvalidSettingsException, CanceledExecutionException, InterruptedException {
        final ExcelSheetWriter sheetWriter =
            createSheetWriter(rowInput.getDataTableSpec(), m_cellWriterFactory, m_cfg.writeRowKey());
        String curSheetName = sheetName;
        var sheetState = getSheet(workbook, curSheetName, sheetWriter);
        var curSheet = sheetState.getFirst();
        long sheetIdx = 0;
        writeColHeaderIfRequired(sheetWriter, curSheet, sheetState.getSecond(), rowInput.getDataTableSpec());
        var rowIdx = sheetWriter.getRowIndex();

        DataRow row;
        while ((row = rowInput.poll()) != null) {
            monitor.checkCanceled();
            while (rowIdx >= m_maxNumRowsPerSheet) {
                // inform the sheet writer that a new sheet is going to come
                finalizeSheet(sheetWriter, curSheet);
                sheetWriter.reset();
                // create the new sheet
                ++sheetIdx;
                curSheetName = SheetUtils.createUniqueSheetName(sheetName, sheetIdx);
                LOGGER.info(String.format("Using additional sheet called '%s' to store entire table", curSheetName));
                sheetState = getSheet(workbook, curSheetName, sheetWriter);
                curSheet = sheetState.getFirst();
                writeColHeaderIfRequired(sheetWriter, curSheet, sheetState.getSecond(), rowInput.getDataTableSpec());
                rowIdx = sheetWriter.getRowIndex();
            }
            monitor.updateProgress(curSheetName, rowIdx);
            sheetWriter.writeRowToSheet(curSheet, row);
            ++rowIdx;
        }
        finalizeSheet(sheetWriter, curSheet);
    }

    /**
     * Creates an instance of {@link ExcelSheetWriter}.
     *
     * @param spec the {@link DataTableSpec}
     * @param cellWriterFactory the {@link ExcelCellWriterFactory}
     * @param writeRowKey flag indicating whether or not to write row keys
     * @return the {@link ExcelSheetWriter}
     */
    abstract ExcelSheetWriter createSheetWriter(final DataTableSpec spec,
        final ExcelCellWriterFactory cellWriterFactory, final boolean writeRowKey);

    private void finalizeSheet(final ExcelSheetWriter sheetWriter, final Sheet curSheet) {
        sheetWriter.setColWidth(curSheet);
        if (m_cfg.useAutoSize()) {
            sheetWriter.autoSizeColumns(curSheet);
        }
    }

    private Pair<Sheet, Boolean> getSheet(final Workbook workbook, final String curSheetName,
        final ExcelSheetWriter sheetWriter) {
        final var pair = getOrCreateSheet(workbook, curSheetName, sheetWriter);
        final var sheet = pair.getFirst();
        if (m_cfg.useAutoSize() && sheet instanceof SXSSFSheet) {
            ((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
        }
        sheet.getPrintSetup().setLandscape(m_cfg.useLandscape());
        sheet.getPrintSetup().setPaperSize(m_cfg.getPaperSize());
        return pair;
    }

    private Pair<Sheet, Boolean> getOrCreateSheet(final Workbook workbook, final String curSheetName,
        final ExcelSheetWriter sheetWriter) {
        final var sheetIdx = workbook.getSheetIndex(curSheetName);
        final Sheet sheet;
        var isNew = sheetIdx == -1;
        final var handling = isNew ? SheetNameExistsHandling.FAIL : m_cfg.getSheetNameExistsHandling();

        switch (handling) {
            case APPEND:
                sheet = prepareSheetAppend(workbook, sheetIdx, sheetWriter);
                break;
            case OVERWRITE:
                isNew = true;
                workbook.removeSheetAt(sheetIdx);
                // fallthrough
            case FAIL:
                sheet = workbook.createSheet(curSheetName);
                break;
            default:
                throw new IllegalStateException("Unexpected SheetNameExistsHandling! " + handling);
        }
        return Pair.create(sheet, isNew);
    }

    private static Sheet prepareSheetAppend(final Workbook workbook, final int sheetIdx,
        final ExcelSheetWriter sheetWriter) {
        final var sheet = workbook.getSheetAt(sheetIdx);
        int lastRow;
        if (sheet instanceof SXSSFSheet) {
            try {
                final var ssheet = (SXSSFSheet)sheet;
                final var field = SXSSFSheet.class.getDeclaredField("_sh");
                field.setAccessible(true); // NOSONAR: I know and I'm sorry but there seems to be no other way
                final var baseSheet = (Sheet)field.get(ssheet);
                lastRow = Math.max(ssheet.getLastRowNum(),
                    Math.max(ssheet.getLastFlushedRowNum(), baseSheet.getLastRowNum())); // this should never be null
            } catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
                throw new IllegalStateException("Could not read last row index", e);
            }
        } else {
            // getLastRowNum returns the index of the last row and not the number
            lastRow = sheet.getLastRowNum();
        }
        sheetWriter.setRowIndex(lastRow + 1);
        return sheet;
    }

    private void writeColHeaderIfRequired(final ExcelSheetWriter sheetWriter, final Sheet sheet, final boolean newSheet,
        final DataTableSpec spec) {
        if (m_cfg.writeColHeaders() && sheetWriter.getRowIndex() < m_maxNumRowsPerSheet
            && (newSheet || !m_cfg.shipColumnHeaderOnAppend())) {
            sheetWriter.writeColHeader(sheet, spec);
            sheet.getLastRowNum();
        }
    }

}
