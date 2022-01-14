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
 *   Nov 13, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.util;

import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Units;

/**
 * Class containing various excel constants.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public final class ExcelConstants {

    /** XLSX file extension */
    public static final String XLSX = ".xlsx";

    /** XLS file extension */
    public static final String XLS = ".xls";

    /** Pattern storing all the characters that are prohibited for sheet name generation by excel. */
    public static final Pattern ILLEGAL_SHEET_NAME_PATTERN = Pattern.compile("[\\[\\]/\\?\\*\\\\:]");

    /** The maximum length of excel sheet names. */
    public static final int MAX_SHEET_NAME_LENGTH = 31;

    /** The maximum column width supported by {@link Sheet#setColumnWidth(int, int)}. */
    public static final int MAX_COL_WIDTH = 255 * 256;

    /** Factor to convert pixels to column width. */
    public static final double PIXEL_TO_COL_WIDTH = MAX_COL_WIDTH * Units.DEFAULT_CHARACTER_WIDTH / Units.EMU_PER_POINT;

    /** xls and xlsx workbook, the maximum image width is 1024 (2^10). */
    public static final int MAX_IMG_WIDTH = 1023;

    /** xls workbooks, the maximum image height is 256 (2^8). */
    public static final int XLS_MAX_IMG_HEIGHT = 255;

    /** The default excel paper size. */
    public static final PaperSize DEFAULT_PAPER_SIZE = PaperSize.A4_PAPERSIZE;

    /** xls workbooks, the row limit is 65,536 rows (2^16). */
    public static final int XLS_MAX_NUM_OF_ROWS = 65536;

    /** xls workbooks, the column limit is 256 columns (2^8). */
    static final int XLS_MAX_NUM_OF_COLS = 256;

    /** xlsx workbooks (and xlsm), the row limit is 1,048,576 rows (2^20). */
    public static final int XLSX_MAX_NUM_OF_ROWS = 1048576;

    /** xlsx workbooks (and xlsm), the column limit is 16,384 columns (2^14). */
    static final int XLSX_MAX_NUM_OF_COLS = 16384;

    private ExcelConstants() {
        // class only containing constants
    }

}
