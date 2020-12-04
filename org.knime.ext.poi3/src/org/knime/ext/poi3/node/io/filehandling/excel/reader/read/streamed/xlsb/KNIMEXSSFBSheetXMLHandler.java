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
 *   Dec 03, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.xlsb;

import java.io.InputStream;
import java.util.BitSet;

import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.XSSFBRecordType;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractKNIMESheetContentsHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractKNIMESheetContentsHandler.KNIMEXSSFDataType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;

/**
 * A slightly modified version of {@link XSSFBSheetHandler} which stores the last cell's type and makes it available to
 * the {@link AbstractKNIMESheetContentsHandler}.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class KNIMEXSSFBSheetXMLHandler extends XSSFBSheetHandler {

    private final AbstractKNIMESheetContentsHandler m_output;

    private KNIMEXSSFDataType m_nextDataType;

    /**
     * @param styles The styles to use.
     * @param strings The {@link String}s.
     * @param sheetContentsHandler The sheet contents.
     * @param dataFormatter Special {@link KNIMEDataFormatter}.
     * @param formulasNotResults Results in formulae instead of their values (usually {@code false}).
     * @see XSSFSheetXMLHandler#XSSFSheetXMLHandler(StylesTable, ReadOnlySharedStringsTable, SheetContentsHandler,
     *      boolean)
     */
    KNIMEXSSFBSheetXMLHandler(final InputStream is, final XSSFBStylesTable styles, final SharedStrings strings,
        final AbstractKNIMESheetContentsHandler sheetContentsHandler, final KNIMEDataFormatter dataFormatter,
        final boolean formulasNotResults) {
        super(is, styles, null, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        m_output = sheetContentsHandler;
    }

    @Override
    public void handleRecord(int id, final byte[] data) { // NOSONAR high Cyclomatic Complexity due to many switch cases
        final XSSFBRecordType type = XSSFBRecordType.lookup(id);
        switch (type) {
            case BrtCellIsst:
            case BrtCellSt:
                m_nextDataType = KNIMEXSSFDataType.STRING;
                break;
            case BrtCellRk:
            case BrtCellReal:
            case BrtFmlaNum:
                m_nextDataType = KNIMEXSSFDataType.NUMBER_OR_DATE;
                break;
            case BrtCellBool:
                m_nextDataType = KNIMEXSSFDataType.BOOLEAN;
                break;
            case BrtFmlaBool:
                m_nextDataType = KNIMEXSSFDataType.BOOLEAN;
                // this record type is not handled by the super class, seems like it simply has been forgotten.
                // we change the id so that it's handled as a regular boolean cell. this should be fine according to
                // the specification(see https://interoperability.blob.core.windows.net/files/MS-XLSB/%5bMS-XLSB%5d.pdf)
                id = XSSFBRecordType.BrtCellBool.getId();
                break;
            case BrtCellError:
            case BrtFmlaError:
                m_nextDataType = KNIMEXSSFDataType.ERROR;
                break;
            case BrtFmlaString:
                m_nextDataType = KNIMEXSSFDataType.FORMULA;
                break;
            case BrtColInfo:
                handleBrtColInfo(data);
                break;
            case BrtRowHdr:
                handleBrtRowHdr(data);
                break;
            default:
                // the others are not of interest for us
                m_nextDataType = null;
                break;
        }
        m_output.nextCellType(m_nextDataType);
        super.handleRecord(id, data);
    }

    private void handleBrtRowHdr(final byte[] data) {
        // according to Microsoft's XLSB specification, the bit specifying whether a row is hidden is at index 92
        // (see https://interoperability.blob.core.windows.net/files/MS-XLSB/%5bMS-XLSB%5d.pdf)
        final BitSet bitSet = BitSet.valueOf(data);
        m_output.hiddenRow(bitSet.size() > 92 && bitSet.get(92));
    }

    private void handleBrtColInfo(final byte[] data) {
        // according to Microsoft's XLSB specification, the bit specifying whether a column is hidden is at index 128
        // (see https://interoperability.blob.core.windows.net/files/MS-XLSB/%5bMS-XLSB%5d.pdf)
        final BitSet bitSet = BitSet.valueOf(data);
        if (bitSet.size() > 128 && bitSet.get(128)) {
            final int colIdx = (int)LittleEndian.getUInt(data, 0);
            m_output.addHiddenCol(colIdx);
        }
    }

}
