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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
 *   23 Aug 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read2;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

/**
 * A slightly modified version of {@link XSSFSheetXMLHandler} which stores the last cell's type.
 *
 * @author Gabor Bakos
 */
@Deprecated
final class KNIMEXSSFSheetXMLHandler extends XSSFSheetXMLHandler {

    /** Data type of the cell. */
    public enum DataType {
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
     * An extension of {@link SheetContentsHandler} that can provide type information of the next cell.
     * @see #nextCellType(DataType)
     */
    interface KNIMESheetContentsHandler extends SheetContentsHandler {
        /**
         * The type of the next cell's type.
         *
         * @param type The type of the next cell.
         */
        void nextCellType(DataType type);

        /**
         * Simple implementation of {@link KNIMESheetContentsHandler}.
         */
        abstract class Abstract implements KNIMESheetContentsHandler {
            protected DataType m_dataType;

            /**
             * {@inheritDoc}
             */
            @Override
            public void startRow(final int rowNum) {
                //Do nothing
            }

            /**
             * {@inheritDoc}
             */
            @Override
            public void endRow(final int rowNum) {
                //Do nothing
            }

            /**
             * {@inheritDoc}
             */
            @Override
            public abstract void cell(final String cellReference, final String formattedValue,
                final XSSFComment comment);

            /**
             * {@inheritDoc}
             */
            @Override
            public void headerFooter(final String text, final boolean isHeader, final String tagName) {
                //Do nothing
            }

            /**
             * {@inheritDoc}
             */
            @Override
            public void nextCellType(final DataType type) {
                m_dataType = type;
            }

        }
    }

    private DataType m_nextDataType;

    private SheetContentsHandler m_output;

    private final Set<Integer> m_hiddenColumns = new LinkedHashSet<>();

    /**
     * @param styles The styles to use.
     * @param strings The {@link String}s.
     * @param sheetContentsHandler The sheet contents.
     * @param dataFormatter Special {@link KNIMEDataFormatter}.
     * @param formulasNotResults Results in formulae instead of their values (usually {@code false}).
     * @see XSSFSheetXMLHandler#XSSFSheetXMLHandler(StylesTable, ReadOnlySharedStringsTable, SheetContentsHandler,
     *      boolean)
     */
    KNIMEXSSFSheetXMLHandler(final StylesTable styles, final ReadOnlySharedStringsTable strings,
        final SheetContentsHandler sheetContentsHandler, final KNIMEDataFormatter dataFormatter,
        final boolean formulasNotResults) {
        super(styles, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        m_output = sheetContentsHandler;
    }

    /**
     * @param styles The styles to use.
     * @param comments The comments.
     * @param strings The {@link String}s.
     * @param sheetContentsHandler The sheet contents.
     * @param dataFormatter Special {@link KNIMEDataFormatter}.
     * @param formulasNotResults Results in formulae instead of their values (usually {@code false}).
     * @see XSSFSheetXMLHandler#XSSFSheetXMLHandler(StylesTable, ReadOnlySharedStringsTable, SheetContentsHandler,
     *      org.apache.poi.ss.usermodel.DataFormatter, boolean)
     */
    KNIMEXSSFSheetXMLHandler(final StylesTable styles, final CommentsTable comments,
        final ReadOnlySharedStringsTable strings, final SheetContentsHandler sheetContentsHandler,
        final KNIMEDataFormatter dataFormatter, final boolean formulasNotResults) {
        super(styles, comments, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        m_output = sheetContentsHandler;
    }

    /**
     * {@inheritDoc} Based on its super implementation. ()
     */
    @Override
    public void startElement(final String uri, final String localName, final String qName, final Attributes attributes)
        throws SAXException {
        if ("c".equals(localName)) {
            // Set up defaults.
            m_nextDataType = DataType.NUMBER_OR_DATE;
            String cellType = attributes.getValue("t");
            if ("b".equals(cellType)) {
                m_nextDataType = DataType.BOOLEAN;
            } else if ("e".equals(cellType)) {
                m_nextDataType = DataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                m_nextDataType = DataType.STRING;
            } else if ("s".equals(cellType)) {
                m_nextDataType = DataType.STRING;
            } else if ("str".equals(cellType)) {
                m_nextDataType = DataType.FORMULA;
            } else {
            }
        }
        if ("col".equals(localName)) {
            String hidden = attributes.getValue("hidden");
            if ("true".equalsIgnoreCase(hidden) || "1".equals(hidden)) {
                try {
                    for (int i = Integer.parseInt(attributes.getValue("min")); i <= Integer
                        .parseInt(attributes.getValue("max")); ++i) {
                        m_hiddenColumns.add(i);
                    }
                } catch (RuntimeException e) {
                    e.printStackTrace();
                }
            }
        }
        super.startElement(uri, localName, qName, attributes);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void endElement(final String uri, final String localName, final String qName) throws SAXException {
        if (m_output instanceof KNIMESheetContentsHandler) {
            ((KNIMESheetContentsHandler)m_output).nextCellType(m_nextDataType);
        }

        super.endElement(uri, localName, qName);
    }

    /**
     * @return the hiddenColumns
     */
    public Set<Integer> getHiddenColumns() {
        return Collections.unmodifiableSet(m_hiddenColumns);
    }
}
