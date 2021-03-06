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
import java.io.InputStream;
import java.nio.file.Path;

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
    public XLSBRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        super(path, config);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    @Override
    public AbstractStreamedParserRunnable createStreamedParser(final InputStream inputStream) throws IOException {
        try {
            m_opc = OPCPackage.open(inputStream);

            final XSSFBReader xssfbReader = new XSSFBReader(m_opc);
            final XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(m_opc);

            m_sheetNames = ExcelUtils.getSheetNames(xssfbReader, sst);
            final SheetIterator sheetsData = (SheetIterator)xssfbReader.getSheetsData();
            m_countingSheetStream = getSheetStreamWithSheetName(sheetsData, getSelectedSheet());
            m_sheetSize = m_countingSheetStream.available();

            // create the parser
            return new XLSBParserRunnable(this, m_config, xssfbReader, sst);
        } catch (SAXException | OpenXML4JException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private class XLSBParserRunnable extends AbstractStreamedParserRunnable {

        private final XSSFBReader m_xssfReader;

        private final SharedStrings m_sharedStringsTable;

        private final KNIMEDataFormatter m_dataFormatter;

        XLSBParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,
            final XSSFBReader xssfReader, final SharedStrings sharedStringsTable) {
            super(read, config);
            m_xssfReader = xssfReader;
            m_sharedStringsTable = sharedStringsTable;
            // Note: Apache POI does not yet support reading out the information whether 1904 windowing us used or not
            // (missing piece is XSSFBRecordType#BrtWbProp). As 1904 is legacy and barely used anymore, we assume false.
            m_dataFormatter = new KNIMEDataFormatter(false, m_use15DigitsPrecision);
        }

        @Override
        protected void parse() throws Exception {
            final ExcelTableReaderSheetContentsHandler sheetContentsHandler =
                new ExcelTableReaderSheetContentsHandler(m_dataFormatter);
            final KNIMEXSSFBSheetXMLHandler sheetHandler = new KNIMEXSSFBSheetXMLHandler(m_countingSheetStream,
                m_xssfReader.getXSSFBStylesTable(), m_sharedStringsTable, sheetContentsHandler, m_dataFormatter, false);
            sheetHandler.parse();
        }

    }

}
