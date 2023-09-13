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
 *   Oct 27, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.xlsx;

import java.util.HashSet;
import java.util.Set;
import java.util.function.Consumer;
import java.util.stream.IntStream;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractKNIMESheetContentsHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractKNIMESheetContentsHandler.KNIMEXSSFDataType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

/**
 * A slightly modified version of {@link XSSFSheetXMLHandler} which stores the last cell's type and makes it available
 * to the {@link AbstractKNIMESheetContentsHandler}.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class KNIMEXSSFSheetXMLHandler extends XSSFSheetXMLHandler {

    private final AbstractKNIMESheetContentsHandler m_output;

    private KNIMEXSSFDataType m_nextDataType;

    private Set<Integer> m_hiddenCols = new HashSet<>();

    private Consumer<Set<Integer>> m_onHiddenColumns;

    /**
     * @param styles The styles to use.
     * @param strings The {@link String}s.
     * @param sheetContentsHandler The sheet contents.
     * @param dataFormatter Special {@link KNIMEDataFormatter}.
     * @param formulasNotResults Results in formulae instead of their values (usually {@code false}).
     * @param onHiddenColumns
     * @see XSSFSheetXMLHandler#XSSFSheetXMLHandler(StylesTable, ReadOnlySharedStringsTable, SheetContentsHandler,
     *      boolean)
     */
    public KNIMEXSSFSheetXMLHandler(final StylesTable styles, final ReadOnlySharedStringsTable strings,
        final AbstractKNIMESheetContentsHandler sheetContentsHandler, final KNIMEDataFormatter dataFormatter,
        final boolean formulasNotResults, final Consumer<Set<Integer>> onHiddenColumns) {
        super(styles, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        m_output = sheetContentsHandler;
        m_onHiddenColumns = onHiddenColumns;
    }

    @Override
    public void startElement(final String uri, final String localName, final String qName, final Attributes attributes)
        throws SAXException {
        // the type detection is based on the super implementation
        if ("c".equals(localName)) {
            // Set up defaults.
            m_nextDataType = KNIMEXSSFDataType.NUMBER_OR_DATE;
            final String cellType = attributes.getValue("t");
            if ("b".equals(cellType)) {
                m_nextDataType = KNIMEXSSFDataType.BOOLEAN;
            } else if ("e".equals(cellType)) {
                m_nextDataType = KNIMEXSSFDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                m_nextDataType = KNIMEXSSFDataType.STRING;
            } else if ("s".equals(cellType)) {
                m_nextDataType = KNIMEXSSFDataType.STRING;
            } else if ("str".equals(cellType)) {
                m_nextDataType = KNIMEXSSFDataType.FORMULA;
            }
        }
        if ("row".equals(localName)) {
            m_output.hiddenRow(isHidden(attributes.getValue("hidden")));
        }
        if ("col".equals(localName) && isHidden(attributes.getValue("hidden"))) {
            int min = Integer.parseInt(attributes.getValue("min"));
            int max = Integer.parseInt(attributes.getValue("max"));
            IntStream.range(min - 1, max).forEach(m_hiddenCols::add);
        }
        super.startElement(uri, localName, qName, attributes);
    }

    private static boolean isHidden(final String hidden) {
        return "true".equalsIgnoreCase(hidden) || "1".equals(hidden);
    }

    @Override
    public void endElement(final String uri, final String localName, final String qName) throws SAXException {
        m_output.nextCellType(m_nextDataType);
        m_output.setHiddenCols(m_hiddenCols);
        if (m_onHiddenColumns != null) {
            m_onHiddenColumns.accept(m_hiddenCols);
        }
        super.endElement(uri, localName, qName);
    }

}
