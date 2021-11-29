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
 *   12.08.2021 (loescher): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.cellcoordinate;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.StringValue;
import org.knime.core.data.xml.io.XMLCellWriterFactory;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.image.ExcelImageWriter;

/**
 * An Excel writer that can update and crate cells at the coordinates given in a table.
 *
 * @author Moditha Hewasinghage, KNIME GmbH, Berlin, Germany
 * @author Jannik Löscher, KNIME GmbH, Konstanz, Germany
 * @see ExcelSheetCellCoordinateWriter#writeCellWithCoordinate(Sheet, DataRow, int, DataTableSpec)
 */
public final class ExcelSheetCellCoordinateWriter extends ExcelSheetWriter {

    /**
     * Constructor.
     *
     * @param spec the specification of the coordinate and column table
     * @param imageWriter the image writer that is used
     * @param cellWriterFactory the cell writer factory used to create writers that update and create cells
     */
    public ExcelSheetCellCoordinateWriter(final DataTableSpec spec, final ExcelImageWriter imageWriter,
        final ExcelCellWriterFactory cellWriterFactory) {
        super(spec, imageWriter, cellWriterFactory, false);
    }

    /**
     * @param sheet the sheet to edit and update
     * @param dataRow a data row specifying the coordinate and value to be written. The index of the coordinate column
     *            is specified by {@code coordinateColumn}. All other columns must be missing values except for at most
     *            one column that contains the value to be written. The written type will be determined from the column
     *            type of the value. If all non-coordinate columns are empty, a missing value will be written to the
     *            coordinate instead. This may be affected by the missing value handling that was configured in the
     *            {@link XMLCellWriterFactory}.
     * @param coordinateColumn the index of the column that contains the coordinates. The coordinate value must be
     *            non-missing string value and will be parsed as a POI cell address.
     * @param spec the table specification of the table containing the coordinates and values. It is used to provide
     *            better error messages.
     * @throws IOException - If the cell could not be written
     * @throws InvalidSettingsException - If there are multiple values to insert found in the row
     */
    public void writeCellWithCoordinate(final Sheet sheet, final DataRow dataRow, final int coordinateColumn,
        final DataTableSpec spec) throws IOException, InvalidSettingsException {

        var idx = 0;
        var found = false;
        CheckUtils.checkSetting(dataRow.getNumCells() > coordinateColumn, "Address column index (%d) does not exits",
            coordinateColumn + 1);
        final var coordinateContainer = dataRow.getCell(coordinateColumn);
        CheckUtils.checkSetting(!coordinateContainer.isMissing(), "Address cell must not be missing!");
        CheckUtils.checkSetting(coordinateContainer instanceof StringValue, "Address cell has to be of type string!");
        final var cellCoordinate = ((StringValue)coordinateContainer).getStringValue();
        for (final DataCell dataCell : dataRow) {
            if (idx != coordinateColumn) {
                if (!dataCell.isMissing() && !found) {
                    writeCellToSheetAtCoordinate(sheet, cellCoordinate, dataCell, idx);
                    found = true;
                } else if (!dataCell.isMissing()) {
                    CheckUtils.checkSetting(false,
                        "Found second non-missing value in row \"%s\" and column \"%s\" (%d) for cell %s to update!",
                        dataRow.getKey(), spec.getColumnSpec(idx).getName(), idx + 1, cellCoordinate);
                }
            }
            idx++;
        }
        if (!found) {
            writeCellToSheetAtCoordinate(sheet, cellCoordinate, DataType.getMissingCell(), coordinateColumn);
        }
    }

    private void writeCellToSheetAtCoordinate(final Sheet sheet, final String cellCoordinate, final DataCell dataCell,
        final int cellWriterIdx) throws IOException, InvalidSettingsException {
        final var ref = parseCellAddress(cellCoordinate);
        final var excelRow = sheet.getRow(ref.getRow());
        final CellStyle currentStyle;
        final Cell cell;
        if (excelRow != null) {
            cell = excelRow.getCell(ref.getColumn());
            if (cell != null) {
                currentStyle = cell.getCellStyle();
                // we remove the cell here to be added fresh later to avoid a bug in POI where corrupted data would
                // be written into a cell when updating a cell of type `inlineStr`
                excelRow.removeCell(cell);
            } else {
                final var rowStyle = excelRow.getRowStyle();

                currentStyle = rowStyle != null ? rowStyle : sheet.getColumnStyle(ref.getColumn());
            }
            writeNewCell(dataCell, cellWriterIdx, ref, excelRow, currentStyle);
        } else {
            currentStyle = sheet.getColumnStyle(ref.getColumn());
            final var newRow = sheet.createRow(ref.getRow());
            writeNewCell(dataCell, cellWriterIdx, ref, newRow, currentStyle);
        }
    }

    private static CellAddress parseCellAddress(final String cellCoordinate) throws InvalidSettingsException {
        final CellAddress ref;
        if (cellCoordinate.indexOf(':') != -1) {
            final int col;
            final int row;
            final var parts = cellCoordinate.split(":");
            CheckUtils.checkSetting(parts.length == 2,
                "Expected cell address \"%s\" to be in \"<column number>:<row number>\" format", cellCoordinate);
            try {
                col = Integer.parseInt(parts[0].trim()) - 1;
            } catch (NumberFormatException e) {
                throw new InvalidSettingsException("Column in '" + cellCoordinate + "' is not an integer!");
            }
            try {
                row = Integer.parseInt(parts[1].trim()) - 1;
            } catch (NumberFormatException e) {
                throw new InvalidSettingsException("Row in '" + cellCoordinate + "' is not an integer!");
            }
            ref = new CellAddress(row, col);
        } else {
            try {
                ref = new CellAddress(cellCoordinate);
            } catch (NumberFormatException e) {
                throw new InvalidSettingsException("Resolved address '" + cellCoordinate + "' is invalid!");
            }
        }
        CheckUtils.checkSetting(ref.getRow() >= 0 && ref.getColumn() >= 0,
            "Resolved address (%s => column %d, row %d) is illegal", cellCoordinate, ref.getColumn() + 1,
            ref.getRow() + 1);
        return ref;

    }

    private void writeNewCell(final DataCell dataCell, final int idx, final CellAddress ref, final Row excelRow,
        final CellStyle style) throws IOException {
        m_cellWriters[idx].write(dataCell, t -> applyStyle(excelRow.createCell(ref.getColumn(), t), style));
    }

    private static Cell applyStyle(final Cell cell, final CellStyle style) {
        cell.setCellStyle(style);
        return cell;
    }
}
