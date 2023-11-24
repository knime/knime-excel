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

import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Consumer;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.eventusermodel.XSSFBReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStrings;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.AbstractStreamedRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.xml.sax.SAXException;

/**
 * This class implements a read that uses the streaming API of Apache POI (eventmodel API) and can read the xlsb file
 * format.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSBRead extends AbstractStreamedRead {

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    public XLSBRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
        final Consumer<Map<String, Boolean>> sheetNamesConsumer) throws IOException {
        super(path, config, sheetNamesConsumer);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @SuppressWarnings("resource") // parser will handle closing sheetStream and pkg
    @Override
    public AbstractStreamedParserRunnable createStreamedParser(final OPCPackage pkg) throws IOException {
        try {
            final var xssfbReader = new XSSFBReader(pkg);
            final var sst = new XSSFBSharedStringsTable(pkg);

            final var sheetNames = ExcelUtils.getSheetNames(xssfbReader, sst);
            if (m_sheetNamesConsumer != null) {
                m_sheetNamesConsumer.accept(sheetNames);
            }
            final SheetIterator sheetsData = (SheetIterator)xssfbReader.getSheetsData();
            final var sheetStream = getSheetStreamWithSheetName(sheetsData, getSelectedSheet(sheetNames));
            // The underlying compressed (zip) input stream always reports "available" as 1 until fully consumed
            // Hence, we get the size from the part (zip entry) itself and use the bytes passed through the counting
            // sheet stream to estimate the progress.
            m_sheetSize = sheetsData.getSheetPart().getSize();

            // create the parser
            return new XLSBParserRunnable(this, m_config, pkg, sheetStream, xssfbReader, sst);
        } catch (SAXException | OpenXML4JException e) {
            throw new IOException(e.getMessage(), e);
        }
    }

    private static class XLSBParserRunnable extends AbstractStreamedParserRunnable {

        private final XSSFBReader m_xssfReader;

        private final SharedStrings m_sharedStringsTable;

        private final KNIMEDataFormatter m_dataFormatter;

        XLSBParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,
                final OPCPackage pkg, final CountingInputStream sheetStream, final XSSFBReader xssfReader,
                final SharedStrings sharedStringsTable) {
            super(read, config, pkg, sheetStream);
            m_xssfReader = xssfReader;
            m_sharedStringsTable = sharedStringsTable;
            // Note: Apache POI does not yet support reading out the information whether 1904 windowing us used or not
            // (missing piece is XSSFBRecordType#BrtWbProp). As 1904 is legacy and barely used anymore, we assume false.
            m_dataFormatter = new KNIMEDataFormatter(false, m_use15DigitsPrecision);
        }

        @Override
        protected void parse() throws Exception {
            final var sheetContentsHandler = new ExcelTableReaderSheetContentsHandler(m_dataFormatter);

            final var sheetHandler = new KNIMEXSSFBSheetXMLHandler(m_sheetStream,
                m_xssfReader.getXSSFBStylesTable(), m_sharedStringsTable, sheetContentsHandler, m_dataFormatter,
                false);
            sheetHandler.parse();

            // XLSB library does not invoke {@link ExcelTableReaderSheetContentsHandler#endSheet()}
            // therefore invoke endSheet() manually.
            sheetContentsHandler.endSheet();
        }
    }
}
