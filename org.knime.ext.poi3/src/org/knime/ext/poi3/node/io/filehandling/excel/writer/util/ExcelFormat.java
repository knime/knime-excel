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
 *   Nov 9, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.util;

import java.util.function.Supplier;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.XlsTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.XlsxTableWriter;

/**
 * The supported excel formats by the 'Excel table writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public enum ExcelFormat {

        /** Type reffering to an xls excel file. */
        XLS(HSSFWorkbook::new, XlsTableWriter::new, ".xls", ExcelConstants.XLS_MAX_NUM_OF_ROWS,
            ExcelConstants.XLS_MAX_NUM_OF_COLS),

        /** Type referring to a xlsx/xlsm excel file. */
        XLSX(SXSSFWorkbook::new, XlsxTableWriter::new, ".xlsx", ExcelConstants.XLSX_MAX_NUM_OF_ROWS,
            ExcelConstants.XLSX_MAX_NUM_OF_COLS);

    private final Supplier<Workbook> m_workbookSupplier;

    private final TableWriterFunction m_createWriter;

    private final String m_fileExtension;

    private final int m_maxRowsPerSheet;

    private final int m_maxColsPerSheet;

    private ExcelFormat(final Supplier<Workbook> workbookSupplier, final TableWriterFunction createWriter,
        final String fileExtension, final int maxRowsPerSheet, final int maxColsPerSheet) {
        m_workbookSupplier = workbookSupplier;
        m_createWriter = createWriter;
        m_fileExtension = fileExtension;
        m_maxRowsPerSheet = maxRowsPerSheet;
        m_maxColsPerSheet = maxColsPerSheet;
    }

    /**
     * Creates the associated {@link ExcelTableWriter} instance.
     *
     * @param cfg the {@link ExcelTableConfig}
     * @param cellWriterFactory the {@link ExcelCellWriterFactory}
     * @return the associated {@link ExcelTableWriter}
     */
    public ExcelTableWriter createWriter(final ExcelTableConfig cfg, final ExcelCellWriterFactory cellWriterFactory) {
        return m_createWriter.createTableWriter(cfg, cellWriterFactory);
    }

    /**
     * Returns the {@link Workbook} associated with this format.
     *
     * @return the {@link Workbook} associated with this format
     */
    public Workbook getWorkbook() {
        return m_workbookSupplier.get();
    }

    /**
     * Returns the file extension associated with this format.
     *
     * @return the file extension associated with this format
     */
    public String getFileExtension() {
        return m_fileExtension;
    }

    /**
     * Returns the maximum number of columns associated with this format.
     *
     * @return the maximum number of columns associated with this format
     */
    public int getMaxNumCols() {
        return m_maxColsPerSheet;
    }

    /**
     * Returns the maximum number of rows per sheet associated with this format.
     *
     * @return the maximum number of rows per sheet associated with this format
     */
    public int getMaxNumRowsPerSheet() {
        return m_maxRowsPerSheet;
    }
}
