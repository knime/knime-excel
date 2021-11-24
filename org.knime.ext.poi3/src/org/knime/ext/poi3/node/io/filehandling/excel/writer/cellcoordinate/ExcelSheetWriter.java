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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.cellcoordinate;

import java.io.IOException;
import java.util.OptionalInt;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Units;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.image.ExcelImageWriter;

/**
 * An excel writer that writes individual rows to a given excel {@link Sheet}.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public class ExcelSheetWriter {

    /** Contains the cell writers used to write cells of differing types. */
    protected final ExcelCellWriter[] m_cellWriters;

    private final boolean m_writeRowKey;

    private final ExcelImageWriter m_imageWriter;

    private int m_rowIdx;

    /**
     * Constructor.
     *
     * @param spec the {@link DataTableSpec}
     * @param imageWriter the {@link ExcelImageWriter}
     * @param cellWriterFactory the {@link ExcelCellWriterFactory}
     * @param writeRowKey flag indicating whether or not to write row keys
     */
    public ExcelSheetWriter(final DataTableSpec spec, final ExcelImageWriter imageWriter,
        final ExcelCellWriterFactory cellWriterFactory, final boolean writeRowKey) {
        m_cellWriters = new ExcelCellWriter[spec.getNumColumns()];
        for (var i = 0; i < m_cellWriters.length; i++) {
            final DataType type = spec.getColumnSpec(i).getType();
            m_cellWriters[i] = cellWriterFactory.createCellWriter(type);
        }
        m_writeRowKey = writeRowKey;
        m_imageWriter = imageWriter;
        m_rowIdx = 0;
    }

    /**
     * Writes the column headers to the given sheet.
     *
     * @param sheet the sheet to add the column header to
     * @param spec the spec containing the column header
     */
    public void writeColHeader(final Sheet sheet, final DataTableSpec spec) {
        final var excelRow = sheet.createRow(m_rowIdx);
        ++m_rowIdx;
        var idx = 0;
        if (m_writeRowKey) {
            excelRow.createCell(idx, CellType.STRING).setCellValue("RowID");
            ++idx;
        }
        for (final DataColumnSpec colSpec : spec) {
            excelRow.createCell(idx, CellType.STRING).setCellValue(colSpec.getName());
            ++idx;
        }
    }

    /**
     * Writes a {@link DataRow} to the given sheet.
     *
     * @param sheet the sheet to write the {@link DataRow} to
     * @param dataRow the {@link DataRow} to be written to the {@link Sheet}
     * @throws IOException - If the {@link DataRow} could not be written to the {@link Sheet}
     */
    public void writeRowToSheet(final Sheet sheet, final DataRow dataRow) throws IOException {
        final var excelRow = sheet.createRow(m_rowIdx);
        var colIdxOffset = 0;
        if (m_writeRowKey) {
            excelRow.createCell(colIdxOffset, CellType.STRING).setCellValue(dataRow.getKey().getString());
            ++colIdxOffset;
        }
        var idx = 0;
        for (final DataCell dataCell : dataRow) {
            final int colIdx = idx + colIdxOffset;
            m_cellWriters[idx].write(dataCell, t -> excelRow.createCell(colIdx, t));
            idx++;
        }
        final OptionalInt rowHeight = m_imageWriter.writeImages(sheet, dataRow, m_rowIdx, colIdxOffset);
        rowHeight.ifPresent(h -> excelRow.setHeightInPoints((float)Units.pixelToPoints(h)));
        ++m_rowIdx;
    }

    /**
     * Resets the writers row index. Must be called before starting to write to a new {@link Sheet}.
     */
    public void reset() {
        m_rowIdx = 0;
        m_imageWriter.reset();
    }

    /**
     * Set the next row which will be written to.
     *
     * @param rowIdx the row index to set
     */
    public void setRowIndex(final int rowIdx) {
        m_rowIdx = rowIdx;
    }

    /**
     * @return the next row which will be written to.
     */
    public int getRowIndex() {
        return m_rowIdx;
    }

    /**
     * Sets the {@link Sheet}'s columns width.
     *
     * @param sheet the {@link Sheet} whose column width must be set
     */
    public void setColWidth(final Sheet sheet) {
        m_imageWriter.setColWidth(sheet, m_writeRowKey ? 1 : 0);
    }

    /**
     * Sets the auto size columns option for the {@link Sheet}.
     *
     * @param sheet the {@link Sheet} whose columns must be auto sized
     */
    public void autoSizeColumns(final Sheet sheet) {
        final int offset = m_writeRowKey ? 1 : 0;
        IntStream.range(0, m_cellWriters.length + offset)//
            .forEach(sheet::autoSizeColumn);
    }
}
