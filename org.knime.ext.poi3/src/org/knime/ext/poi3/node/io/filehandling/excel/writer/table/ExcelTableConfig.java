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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.table;

import java.util.Optional;

import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetNameExistsHandling;

/**
 * The excel table configuration interface.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public interface ExcelTableConfig {

    /**
     * Returns the sheet names.
     *
     * @return the sheet names
     */
    String[] getSheetNames();

    /**
     * Returns the missing value pattern if selected.
     *
     * @return the missing value pattern if selected
     */
    Optional<String> getMissingValPattern();

    /**
     * Flag indicating whether or not the sheet's columns have to be auto sized.
     *
     * @return {@code true} if the sheet's columns have to be auto sized and {@code false} otherwise
     */
    boolean useAutoSize();

    /**
     * Flag indicating whether or not the sheet uses the landscape layout.
     *
     * @return {@code true} if the sheet's layout has to be set to landscape and {@code false} otherwise
     */
    boolean useLandscape();

    /**
     * Returns the paper size.
     *
     * @return the paper size
     */
    short getPaperSize();

    /**
     * Flag indicating whether or not row keys must be written to the sheet.
     *
     * @return {@code true} if the row keys must be written to the sheet and {@code false} otherwise
     */
    boolean writeRowKey();

    /**
     * Flag indicating whether or not column headers must be written to the sheet.
     *
     * @return {@code true} if the column headers must be written to the sheet and {@code false} otherwise
     */
    boolean writeColHeaders();

    /**
     * Flag indicating whether writing of column headers should be skipped if the sheet already exists.
     * This option only has an effect if {@link #writeColHeaders()} is {@code true}
     *
     * @return {@code true} if the column headers should be skipped if the sheet already exists and {@code false} otherwise
     */
    default boolean shipColumnHeaderOnAppend() {
        return true;
    }

    /**
     * Flag indicating whether or not overwriting sheets is allowed.
     *
     * @return {@code true} overwriting sheets is forbidden and the execution must be stopped
     * @nooverride implement or use {@link #getSheetNameExistsHandling()} instead
     */
    default boolean abortIfSheetExists() {
        return getSheetNameExistsHandling() == SheetNameExistsHandling.FAIL;
    }

    /**
     * Returns how sheets that already exist should be handled.
     *
     * @return {@code true} overwriting sheets is forbidden and the execution must be stopped
     */
    default SheetNameExistsHandling getSheetNameExistsHandling() {
        return abortIfSheetExists() ? SheetNameExistsHandling.FAIL : SheetNameExistsHandling.OVERWRITE;
    }

    /**
     * Flag indicating whether or not formulas have to be (re-)evaluated after all sheets have been written.
     *
     * @return {@code true} if the sheets need to be (re-)evaluated and {@code false} otherwise
     */
    boolean evaluate();

    /**
     *
     * Returns the {@link ExcelFormat} to be written.
     *
     * @return the {@link ExcelFormat} to be written
     */
    ExcelFormat getExcelFormat();
}