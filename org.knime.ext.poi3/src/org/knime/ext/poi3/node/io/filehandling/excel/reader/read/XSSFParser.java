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
 *   14 Sep 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.function.Consumer;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.CellConsumer;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.SheetMetadata;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.WorkbookMetadata;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.ExcelTableReaderSheetContentsHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.KNIMEDataFormatter;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed.xlsx.KNIMEXSSFSheetXMLHandler;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
final class XSSFParser implements XLParser {

    private CountingInputStream m_countingSheetStream;
    private OPCPackage m_pkg;

    private final XSSFReader m_xssfReader;
    private final XMLReader m_xmlReader;
    private final ReadOnlySharedStringsTable m_sharedStringsTable;

    private TableReadConfig<ExcelTableReaderConfig> m_config;
    private CompletableFuture<WorkbookMetadata> m_metadata;
    private CellConsumer m_output;

    private Path m_path;

    private XSSFParser(final OPCPackage pkg, final TableReadConfig<ExcelTableReaderConfig> config,
            final CompletableFuture<WorkbookMetadata> metadata, final CellConsumer output, final Path path)
            throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
        m_pkg = pkg;
        m_metadata = metadata;
        m_output = output;
        m_path = path;
        m_xssfReader = new XSSFReader(pkg);
        m_xmlReader = SAXHelper.newXMLReader();
        // disable DTD to prevent almost all XXE attacks, XMLHelper.newXMLReader() did set further security features
        m_xmlReader.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        m_sharedStringsTable = new ReadOnlySharedStringsTable(pkg, false);

        m_config = config;
    }

    static XSSFParser createParser(final OPCPackage pkg, final TableReadConfig<ExcelTableReaderConfig> config,
            final CompletableFuture<WorkbookMetadata> metadata, final CellConsumer output, final Path path)
            throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
        if (isBinaryFile(pkg)) {
            throw new UnsupportedOperationException("Not yet implemented!");
        }
        return new XSSFParser(pkg, config, metadata, output, path);
    }

    private static boolean isBinaryFile(final OPCPackage pkg) {
        var core = pkg.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
        if (core.size() == 0) {
            core = pkg.getRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT);
        }
        final var corePart = pkg.getPart(core.getRelationship(0));
        final var coreType = corePart.getContentType();
        // use binary parser, else try xml
        return XSSFRelation.XLSB_BINARY_WORKBOOK.getContentType().equals(coreType);
    }

    @Override
    public void close() throws IOException {
        m_pkg.close();
    }

    @Override
    public void parse() throws InvalidFormatException, IOException {
        try {
            final var sheetNames = getSheetNames();
            final var sheetSize = openSelectedSheet();

            final var dataFormatter = new KNIMEDataFormatter(use1904Windowing(m_xssfReader),
                m_config.getReaderSpecificConfig().isUse15DigitsPrecision());

            final var sheetMeta = new CompletableFuture<SheetMetadata>();

            final Consumer<Set<Integer>> onHiddenColumns  = hiddenColumns ->
                sheetMeta.complete(new SheetMetadata(sheetSize, hiddenColumns));

            sheetMeta.whenComplete((sheet, ex) -> {
                if (ex != null) {
                    m_metadata.completeExceptionally(ex);
                } else {
                    m_metadata.complete(new WorkbookMetadata(sheetNames, sheet));
                }
            });

            final var sheetContentsHandler = new ExcelTableReaderSheetContentsHandler(m_output, m_config,
                dataFormatter, this::getProgress);
            m_xmlReader.setContentHandler(new KNIMEXSSFSheetXMLHandler(m_xssfReader.getStylesTable(),
                m_sharedStringsTable, sheetContentsHandler, dataFormatter, false, onHiddenColumns));
            m_xmlReader.parse(new InputSource(m_countingSheetStream));
        } catch (final SAXException e) {
            throw new IOException(e);
        }
    }

    private Map<String, Boolean> getSheetNames() throws InvalidFormatException, IOException, SAXException {
        return ExcelUtils.getSheetNames(m_xmlReader, m_xssfReader, m_sharedStringsTable);
    }

    private static boolean use1904Windowing(final XSSFReader xssfReader) {
        try (final InputStream workbookXml = xssfReader.getWorkbookData()) {
            final WorkbookDocument doc = WorkbookDocument.Factory.parse(workbookXml);
            final CTWorkbook wb = doc.getWorkbook();
            final CTWorkbookPr prefix = wb.getWorkbookPr();
            return prefix.getDate1904();
        } catch (Exception e) { // NOSONAR
            // if anything goes wrong, we just assume false
            return false;
        }
    }

    long getProgress() {
        if (m_countingSheetStream == null) {
            throw new IllegalStateException("Sheet stream is not open!");
        }
        return m_countingSheetStream.getByteCount();
    }

    private long openSelectedSheet() throws InvalidFormatException, IOException, SAXException {
        if (m_countingSheetStream != null) {
            throw new IllegalStateException("Stream of selected sheet is already open!");
        }
        final var sheetNames = getSheetNames();
        final ExcelTableReaderConfig excelConfig = m_config.getReaderSpecificConfig();
        final var sheet = excelConfig.getSheetSelection().getSelectedSheet(sheetNames, excelConfig, m_path);
        final SheetIterator sheetsData = (SheetIterator)m_xssfReader.getSheetsData();
        m_countingSheetStream = getSheetStreamByName(sheetsData, sheet);
        return sheetsData.getSheetPart().getSize();
    }

    private static CountingInputStream getSheetStreamByName(final SheetIterator sheetsData, final String sheet)
            throws IOException {
        while (sheetsData.hasNext()) {
            @SuppressWarnings("resource") // manually closed
            final InputStream sheetIterator = sheetsData.next();
            try {
                if (sheetsData.getSheetName().equals(sheet)) {
                    return new CountingInputStream(sheetIterator);
                } else {
                    sheetIterator.close();
                }
            } catch (Exception e) { // NOSONAR catch anything to make sure the stream is closed
                sheetIterator.close();
            }
        }
        // should never happen as we checked for it already
        throw new IllegalStateException("No sheet with name \"" + sheet + "\" found.");
    }

}
