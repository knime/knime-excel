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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.knime.filehandling.core.node.table.reader.config.ReaderSpecificConfig;

/**
 * An implementation of {@link ReaderSpecificConfig} class for Excel table reader.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelTableReaderConfig implements ReaderSpecificConfig<ExcelTableReaderConfig> {

    private boolean m_use15DigitsPrecision = true;

    private SheetSelection m_sheetSelection = SheetSelection.FIRST;

    private String m_sheetName = "";

    private int m_sheetIdx = 0;

    private boolean m_skipHiddenCols;

    private boolean m_skipHiddenRows;

    /**
     * Constructor.
     */
    ExcelTableReaderConfig() {
    }

    private ExcelTableReaderConfig(final ExcelTableReaderConfig toCopy) {
        setUse15DigitsPrecision(toCopy.isUse15DigitsPrecision());
        setSheetSelection(toCopy.getSheetSelection());
        setSheetName(toCopy.getSheetName());
        setSheetIdx(toCopy.getSheetIdx());
        setSkipHiddenCols(toCopy.isSkipHiddenCols());
        setSkipHiddenRows(toCopy.isSkipHiddenRows());
    }

    @Override
    public ExcelTableReaderConfig copy() {
        return new ExcelTableReaderConfig(this);
    }

    /**
     * @return the use15DigitsPrecision
     */
    public boolean isUse15DigitsPrecision() {
        return m_use15DigitsPrecision;
    }

    /**
     * @param use15DigitsPrecision the use15DigitsPrecision to set
     */
    void setUse15DigitsPrecision(final boolean use15DigitsPrecision) {
        m_use15DigitsPrecision = use15DigitsPrecision;
    }

    /**
     * @return the sheetSelection
     */
    public SheetSelection getSheetSelection() {
        return m_sheetSelection;
    }

    /**
     * @param sheetSelection the sheetSelection to set
     */
    void setSheetSelection(final SheetSelection sheetSelection) {
        m_sheetSelection = sheetSelection;
    }

    /**
     * @return the selectedSheetName
     */
    public String getSheetName() {
        return m_sheetName;
    }

    /**
     * @param selectedSheetName the selectedSheetName to set
     */
    void setSheetName(final String selectedSheetName) {
        m_sheetName = selectedSheetName;
    }

    /**
     * @return the selectedSheetIdx
     */
    public int getSheetIdx() {
        return m_sheetIdx;
    }

    /**
     * @param selectedSheetIdx the selectedSheetIdx to set
     */
    void setSheetIdx(final int selectedSheetIdx) {
        m_sheetIdx = selectedSheetIdx;
    }

    /**
     * @return the skipHiddenCols
     */
    public boolean isSkipHiddenCols() {
        return m_skipHiddenCols;
    }

    /**
     * @param skipHiddenCols the skipHiddenCols to set
     */
    public void setSkipHiddenCols(final boolean skipHiddenCols) {
        m_skipHiddenCols = skipHiddenCols;
    }

    /**
     * @return the skipHiddenRows
     */
    public boolean isSkipHiddenRows() {
        return m_skipHiddenRows;
    }

    /**
     * @param skipHiddenRows the skipHiddenRows to set
     */
    public void setSkipHiddenRows(final boolean skipHiddenRows) {
        m_skipHiddenRows = skipHiddenRows;
    }

}