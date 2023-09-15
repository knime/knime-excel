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

import java.io.BufferedInputStream;
import java.io.IOException;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.concurrent.CompletableFuture;
import java.util.function.Consumer;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipFileZipEntrySource;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.CellConsumer;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead2.WorkbookMetadata;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils.NotExcelFileException;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils.ParsingInterruptedException;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.xml.sax.SAXException;

/**
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public class ExcelParser implements Runnable {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelParser.class);
    private final Path m_path;
    private final TableReadConfig<ExcelTableReaderConfig> m_config;
    private final CellConsumer m_output;

    private final Consumer<Throwable> m_exceptionHandler;
    private CompletableFuture<WorkbookMetadata> m_metadata;

    public ExcelParser(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
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
                final var parser = createParser(channel)) {

            // get sheet names
            // get selected sheet name
            // get sheet size -> max progress
            // parse rows
            // -- get hidden columns
            parser.parse();
        } catch (final Exception e) {
            // that one would be us
            if (!(e instanceof ParsingInterruptedException)) {
                m_exceptionHandler.accept(e);
            }
            m_metadata.completeExceptionally(e);
        } finally {
            try {
                m_output.finished();
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                // TODO handle?
            }
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
     * @throws GeneralSecurityException
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    private XLParser createParser(final SeekableByteChannel ch) throws IOException, GeneralSecurityException,
            OpenXML4JException, SAXException, ParserConfigurationException {
        // identify file type and open package/workbook
        final var fileType = ExcelUtils.sniffFileType(ch);
        switch (fileType) {
            case OLE2 -> {
                final var asm = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
                final var authType = asm.getAuthenticationType();
                final var password = authType == AuthenticationType.NONE ? null
                    : asm.getPassword(m_config.getReaderSpecificConfig().getCredentialsProvider());
                return createParserFromOLE2(ch, password);
            }
            case OOXML -> {
                return XSSFParser.createParser(OPCPackage.open(new ZipFileZipEntrySource(new ZipFile(ch))), m_config,
                    m_metadata, m_output, m_path);
            }
            default -> throw new NotExcelFileException(
                "File format \"%s\" is not supported for spreadsheet reading.".formatted(fileType));
        }
    }

    /**
     * Creates a parser on the given channel containing a valid OLE2 file.
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    private XLParser createParserFromOLE2(final SeekableByteChannel channel, final String password)
            throws IOException, GeneralSecurityException, OpenXML4JException, SAXException, ParserConfigurationException {
        try (final var fs = openFS(channel)) {
            // first try for XML-based files
            if (fs.getRoot().hasEntry(Decryptor.DEFAULT_POIFS_ENTRY)) {
                // encrypted XLSX/XLSB
                final var ei = new EncryptionInfo(fs); // this will fail on unencrypted files!
                final var d = Decryptor.getInstance(ei);
                if (!d.verifyPassword(password)) {
                    throw new EncryptedDocumentException("Supplied password is invalid");
                }
                // encrypted files will be buffered in memory fully, since we have to use
                // OPCPackage.open(InputStream), because Decryptor only operates on streams (and not channels)
                // Using `AesZipFileZipEntrySource.createZipEntrySource(decryptedStream)` to buffer the file on
                // disk fails with "Truncated ZIP file"
                return XSSFParser.createParser(OPCPackage.open(d.getDataStream(fs)), m_config, m_metadata, m_output,
                    m_path);
            } else {
                // then test XLS
                return new SSFParser(WorkbookFactory.create(fs.getRoot(), password)); // will hold whole sheet in memory
            }
        }
    }

    @SuppressWarnings("resource")
    private static POIFSFileSystem openFS(final SeekableByteChannel channel) throws IOException {
        if (channel instanceof FileChannel fchan) {
            return new POIFSFileSystem(fchan);
        }
        // POIFSFileSystem closes the stream after copying the contents
        return new POIFSFileSystem(new BufferedInputStream(Channels.newInputStream(channel)));
    }
}
