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
 *   Dec 3, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed;

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Abstract class implementing {@link SheetContentsHandler} that stores information extracted by sheet handlers during
 * parsing.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class AbstractKNIMESheetContentsHandler implements SheetContentsHandler {

    private KNIMEXSSFDataType m_dataType;

    private Set<Integer> m_hiddenCols = new HashSet<>();

    private boolean m_isHiddenRow;

    /**
     * Used for recognizing empty <c> elements, for which
     * {@link SheetContentsHandler#cell(String, String, org.apache.poi.xssf.usermodel.XSSFComment)} is never called.
     * In case it is called, the current implementation will reset the flag. (AP-21960)
     */
    private boolean m_expectCellContent;

    private String m_expectCellReference;

    /**
     * @param type the {@link KNIMEXSSFDataType} of the cell
     */
    public void nextCellType(final KNIMEXSSFDataType type) {
        m_dataType = type;
    }

    KNIMEXSSFDataType getCellType() {
        return m_dataType;
    }

    /**
     * @param hiddenCols the hidden column indices
     */
    public void setHiddenCols(final Set<Integer> hiddenCols) {
        m_hiddenCols = hiddenCols;
    }

    Set<Integer> getHiddenCols() {
        return m_hiddenCols;
    }

    /**
     * @param isHiddenRow whether the row is hidden
     */
    public void hiddenRow(final boolean isHiddenRow) {
        m_isHiddenRow = isHiddenRow;
    }

    boolean isHiddenRow() {
        return m_isHiddenRow;
    }

    /** Data type of the cell. */
    public enum KNIMEXSSFDataType {
            /** Number or date */
            NUMBER_OR_DATE,
            /** Boolean */
            BOOLEAN,
            /** Error in the formula */
            ERROR,
            /** Text */
            STRING,
            /** A formula */
            FORMULA;
    }

    /**
     * Sets the property that a cell element ('c' tag) is to be expected, required to handle sparse rows, see AP-21960.
     * @param cellReference The cell reference = 'r' attribute, e.g. "A6", can be null (order defines reference).
     */
    public void setExpectsCellContent(final String cellReference) {
        m_expectCellContent = true;
        m_expectCellReference = cellReference;
    }

    /**
     * @return true if in a 'c' tag, and haven't seen content, see {@link #setExpectsCellContent(String)}
     */
    public boolean expectsCellContent() {
        return m_expectCellContent;
    }

    /**
     * @return Get value set in {@link #setExpectsCellContent(String)}
     */
    public String getExpectedCellReference() {
        return m_expectCellReference;
    }

    @Override
    public void cell(final String cellReference, final String formattedValue, final XSSFComment comment) {
        m_expectCellContent = false;
        m_expectCellReference = null;
    }

    /**
     * Called by the parser in case an empty cell was visited, addresses AP-21960.
     */
    public abstract void handleEmptyCell();

}
