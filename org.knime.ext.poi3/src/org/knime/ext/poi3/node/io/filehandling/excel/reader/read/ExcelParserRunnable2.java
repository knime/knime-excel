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
 *   12 Sep 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.io.InputStream;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipFileZipEntrySource;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.CellConsumer;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.SheetMetadata;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.WorkbookMetadata;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils.ParsingInterruptedException;
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
public class ExcelParserRunnable2 implements Runnable {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelParserRunnable2.class);
    private final Path m_path;
    private final TableReadConfig<ExcelTableReaderConfig> m_config;
    private final CellConsumer m_output;

    private final Consumer<Throwable> m_exceptionHandler;
    private CompletableFuture<WorkbookMetadata> m_metadata;

    public ExcelParserRunnable2(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
            final CellConsumer output, final CompletableFuture<WorkbookMetadata> metadata,
            final Consumer<Throwable> exceptionHandler) {
        m_path = path;
        m_config = config;
        m_output = output;
        m_metadata = metadata;
        m_exceptionHandler = exceptionHandler;
    }


    @Override
    public void run() {
        LOGGER.debug("Thread parsing Excel spreadsheet started.");
        try (final var channel = Files.newByteChannel(m_path);
                final var pkg = openPackage(channel);
                final var parser = createParser(pkg)) {

            // get sheet names
            // get selected sheet name
            // get sheet size -> max progress
            // parse rows
            // -- get hidden columns
            parser.parse();
        } catch (final IOException | OpenXML4JException | SAXException | ParserConfigurationException |
                GeneralSecurityException e) {
            m_exceptionHandler.accept(e);
        } catch (ParsingInterruptedException e) {
            // that's us
        } finally {
            try {
                m_output.finished();
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                // TODO handle?
            }
        }
    }

    private XLSXParser createParser(final OPCPackage pkg) throws IOException, OpenXML4JException, SAXException,
            ParserConfigurationException {
        return new XLSXParser(pkg);
    }

    /**
     * Not thread-safe!
     *
     * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
     */
    private final class XLSXParser implements AutoCloseable {

        private CountingInputStream m_countingSheetStream;

        private final XSSFReader m_xssfReader;
        private final XMLReader m_xmlReader;
        private final ReadOnlySharedStringsTable m_sharedStringsTable;


        private XLSXParser(final OPCPackage pkg) throws IOException, OpenXML4JException, SAXException,
                ParserConfigurationException {
            m_xssfReader = new XSSFReader(pkg);
            m_xmlReader = SAXHelper.newXMLReader();
            // disable DTD to prevent almost all XXE attacks, XMLHelper.newXMLReader() did set further security features
            m_xmlReader.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            m_sharedStringsTable = new ReadOnlySharedStringsTable(pkg, false);
        }

        private Map<String, Boolean> getSheetNames() throws InvalidFormatException, IOException, SAXException {
            return ExcelUtils.getSheetNames(m_xmlReader, m_xssfReader, m_sharedStringsTable);
        }

        /**
         * @throws SAXException
         * @throws IOException
         * @throws InvalidFormatException
         *
         */
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

        void parse() throws InvalidFormatException, IOException, SAXException {
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

        @Override
        public void close() throws IOException {
            m_countingSheetStream.close();
        }

        long getProgress() {
            if (m_countingSheetStream == null) {
                throw new IllegalStateException("Sheet stream is not open!");
            }
            return m_countingSheetStream.getByteCount();
        }

    }


    /**
     * Creates and returns an {@link IOException} with an error message telling the user which file requires a password
     * to be opened.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    private IOException createPasswordProtectedFileException(final EncryptedDocumentException e) {
        return new IOException(String
            .format("The file '%s' is password protected. Supply a password via the encryption settings.", m_path), e);
    }

    /**
     * Creates and returns an {@link IOException} with an error message telling the user for which file the password is
     * incorrect.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    private IOException createPasswordIncorrectException(final EncryptedDocumentException e) {
        return new IOException(String.format("The supplied password is incorrect for file '%s'.", m_path), e);
    }

    /**
     * @param channel
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     * @throws GeneralSecurityException
     */
    private OPCPackage openPackage(final SeekableByteChannel ch) throws IOException, InvalidFormatException, GeneralSecurityException {
        // identify file type and open package/workbook
        final var fileType = ExcelUtils.sniffFileType(ch);
        switch (fileType) {
            case OLE2 -> {
                 // OLE2 file could contain:
                // - encrypted XML content (OOXML)
                // - XLS file with XLSX extension
                final var asm = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
                final var authType = asm.getAuthenticationType();
                if (authType == AuthenticationType.NONE) {
                    // This exception would previously be thrown by OPCPackage.open(File),
                    // but is not by OPCPackage.open(InputStream), so we have to throw it ourselves!
                    // The ExcelTableReader catches this and switches to the XLSRead
                    throw new OLE2NotOfficeXmlFileException("OLE2 file but trying to open as OOXML-based file.");
                }
                final var password = asm.getPassword(m_config.getReaderSpecificConfig().getCredentialsProvider());
                return openDecrypting(ch, password, () -> createPasswordIncorrectException(null));
            }
            case OOXML -> {
                return OPCPackage.open(new ZipFileZipEntrySource(new ZipFile(ch)));
            }
            default -> throw new NotOfficeXmlFileException("");
        }
    }

    private static OPCPackage openDecrypting(final SeekableByteChannel chan, final String password,
            final Supplier<IOException> passwordIncorrect) throws IOException, GeneralSecurityException,
            InvalidFormatException {
        try (final var poiFS = newPOIFS(chan)) {
            final var encryptionInfo = new EncryptionInfo(poiFS);
            final var decryptor = Decryptor.getInstance(encryptionInfo);
            if (!decryptor.verifyPassword(password)) {
                throw passwordIncorrect.get();
            }
            try (final var decryptedStream = decryptor.getDataStream(poiFS)) {
                // encrypted files will be buffered in memory fully, since we have to use
                // OPCPackage.open(InputStream), because Decryptor only operates on streams (and not channels)
                // Using `AesZipFileZipEntrySource.createZipEntrySource(decryptedStream)` to buffer the file on
                // disk fails with "Truncated ZIP file"
                return OPCPackage.open(decryptedStream);
            }
        }
    }

    private static POIFSFileSystem newPOIFS(final SeekableByteChannel chan) throws IOException {
        if (chan instanceof FileChannel fchan) {
            return new POIFSFileSystem(fchan);
        }
        try (final var in = Channels.newInputStream(chan)) {
            return new POIFSFileSystem(in);
        }
    }
}
