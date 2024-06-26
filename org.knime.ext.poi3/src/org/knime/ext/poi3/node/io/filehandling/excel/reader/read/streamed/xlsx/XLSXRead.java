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
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Consumer;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.xmlbeans.XmlException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
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
     * @param sheetNamesConsumer
     * @throws IOException if an I/O exception occurs
     */
    public XLSXRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
        final Consumer<Map<String, Boolean>> sheetNamesConsumer) throws IOException {
        super(path, config, sheetNamesConsumer);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @SuppressWarnings("resource") // parser will handle closing sheetStream and pkg
    @Override
    public AbstractStreamedParserRunnable createStreamedParser(final OPCPackage pkg) throws IOException {
        try {
            final var xssfReader = new XSSFReader(pkg);
            final var xmlReader = XMLHelper.newXMLReader();
            // disable DTD to prevent almost all XXE attacks, XMLHelper.newXMLReader() did set further security features
            xmlReader.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            final var sharedStringsTable = new ReadOnlySharedStringsTable(pkg, false);

            final var sheetNames = ExcelUtils.getSheetNames(xmlReader, xssfReader, sharedStringsTable);
            if (m_sheetNamesConsumer != null) {
                m_sheetNamesConsumer.accept(sheetNames);
            }

            final SheetIterator sheetsData = (SheetIterator)xssfReader.getSheetsData();
            final var sheetStream = getSheetStreamWithSheetName(sheetsData, getSelectedSheet(sheetNames));
            // The underlying compressed (zip) input stream (InflaterInputStream) always reports "available" as 1
            // until fully consumed, in which case it returns 0
            // Hence, we get the size from the part (zip entry) itself and use the bytes passed through the counting
            // sheet stream to estimate the progress.
            m_sheetSize = sheetsData.getSheetPart().getSize();

            return new XLSXParserRunnable(this, m_config, pkg, sheetStream, xmlReader, xssfReader,
                sharedStringsTable, use1904Windowing(xssfReader));
        } catch (SAXException | XmlException | OpenXML4JException | ParserConfigurationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private static boolean use1904Windowing(final XSSFReader xssfReader)
            throws XmlException, IOException, InvalidFormatException {
        try (final var workbookXml = xssfReader.getWorkbookData()) {
            // this does not parse actual sheet data, just the workbook information
            final var doc = WorkbookDocument.Factory.parse(workbookXml, POIXMLTypeLoader.DEFAULT_XML_OPTIONS);
            final var wb = doc.getWorkbook();
            final var prefix = wb.getWorkbookPr();
            // same as in POI's `XSSFWorkbook#isDate1904()`
            return prefix != null && prefix.getDate1904();
        }
    }

    private static class XLSXParserRunnable extends AbstractStreamedParserRunnable {

        private final XSSFReader m_xssfReader;

        private final ReadOnlySharedStringsTable m_sharedStringsTable;

        private final KNIMEDataFormatter m_dataFormatter;

        private final XMLReader m_xmlReader;

        XLSXParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,  // NOSONAR
                final OPCPackage pkg, final CountingInputStream sheetStream,
                final XMLReader xmlReader, final XSSFReader xssfReader,
                final ReadOnlySharedStringsTable sharedStringsTable, final boolean use1904Windowing) {
            super(read, config, pkg, sheetStream);
            m_xmlReader = xmlReader;
            m_xssfReader = xssfReader;
            m_sharedStringsTable = sharedStringsTable;
            m_dataFormatter = new KNIMEDataFormatter(use1904Windowing, m_use15DigitsPrecision);
        }

        @Override
        protected void parse() throws Exception {
            final var sheetContentsHandler = new ExcelTableReaderSheetContentsHandler(m_dataFormatter);
            m_xmlReader.setContentHandler(new KNIMEXSSFSheetXMLHandler(m_xssfReader.getStylesTable(),
                m_sharedStringsTable, sheetContentsHandler, m_dataFormatter, false));
            m_xmlReader.parse(new InputSource(m_sheetStream));
        }

    }
}
