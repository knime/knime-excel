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
 *   Oct 20, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.xlsx;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.xmlbeans.XmlObject;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * This class implements a read that uses the streaming API of Apache POI (eventmodel API) and can read xlsx and xlsm
 * file formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSXRead extends AbstractStreamedRead {

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    public XLSXRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        super(path, config);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @Override
    public AbstractStreamedParserRunnable createStreamedParser(final InputStream inputStream) throws IOException {
        try {
            m_opc = OPCPackage.open(inputStream);
            final XSSFReader xssfReader = new XSSFReader(m_opc);
            final XMLReader xmlReader = XMLHelper.newXMLReader();
            // disable DTD to prevent almost all XXE attacks, XMLHelper.newXMLReader() did set further security features
            xmlReader.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            final ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(m_opc, false);

            m_sheetNames = ExcelUtils.getSheetNames(xmlReader, xssfReader, sharedStringsTable);
            final SheetIterator sheetsData = (SheetIterator)xssfReader.getSheetsData();
            m_countingSheetStream = getSheetStreamWithSheetName(sheetsData, getSelectedSheet());
            m_sheetSize = m_countingSheetStream.available();

            // create the parser
            return new XLSXParserRunnable(this, m_config, xmlReader, xssfReader, sharedStringsTable,
                use1904Windowing(xssfReader));
        } catch (SAXException | OpenXML4JException | ParserConfigurationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private static boolean use1904Windowing(final XSSFReader xssfReader) {
        try (final InputStream workbookXml = xssfReader.getWorkbookData()) {
            final WorkbookDocument doc = (WorkbookDocument)XmlObject.Factory.parse(workbookXml);
            final CTWorkbook wb = doc.getWorkbook();
            final CTWorkbookPr prefix = wb.getWorkbookPr();
            return prefix.getDate1904();
        } catch (Exception e) { // NOSONAR
            // if anything goes wrong, we just assume false
            return false;
        }
    }

    private class XLSXParserRunnable extends AbstractStreamedParserRunnable {

        private final XSSFReader m_xssfReader;

        private final ReadOnlySharedStringsTable m_sharedStringsTable;

        private final KNIMEDataFormatter m_dataFormatter;

        private final XMLReader m_xmlReader;

        XLSXParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,
            final XMLReader xmlReader, final XSSFReader xssfReader, final ReadOnlySharedStringsTable sharedStringsTable,
            final boolean use1904Windowing) {
            super(read, config);
            m_xmlReader = xmlReader;
            m_xssfReader = xssfReader;
            m_sharedStringsTable = sharedStringsTable;
            m_dataFormatter = new KNIMEDataFormatter(use1904Windowing, m_use15DigitsPrecision);
        }

        @Override
        protected void parse() throws Exception {
            final ExcelTableReaderSheetContentsHandler sheetContentsHandler =
                new ExcelTableReaderSheetContentsHandler(m_dataFormatter);
            m_xmlReader.setContentHandler(new KNIMEXSSFSheetXMLHandler(m_xssfReader.getStylesTable(),
                m_sharedStringsTable, sheetContentsHandler, m_dataFormatter, false));
            m_xmlReader.parse(new InputSource(m_countingSheetStream));
        }

    }
}
