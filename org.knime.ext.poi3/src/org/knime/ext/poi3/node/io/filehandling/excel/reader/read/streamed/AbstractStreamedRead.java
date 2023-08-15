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
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed;

import java.io.IOException;
import java.io.InputStream;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.Map;
import java.util.OptionalLong;
import java.util.Set;
import java.util.function.Supplier;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.FSZipPackage;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;

/**
 * This class implements a read that uses the streaming API of Apache POI (eventmodel API) and can read xlsx and xlsm
 * file formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class AbstractStreamedRead extends ExcelRead {

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    protected AbstractStreamedRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config)
        throws IOException {
        super(path, config);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    /** The sheet stream. */
    protected CountingInputStream m_countingSheetStream;

    /** The package where the sheet stream originates from. */
    private OPCPackage m_opc;

    /** The channel from which the OPCPackage is opened. */
    private SeekableByteChannel m_channel;

    /** The size of the sheet to read. */
    protected long m_sheetSize;

    /** The map of sheet names with names as keys being in the same order as in the workbook. */
    protected Map<String, Boolean> m_sheetNames;

    private AbstractStreamedParserRunnable m_streamedParser;



    @Override
    protected ExcelParserRunnable createParser(final Path path) throws IOException {
        m_channel = Files.newByteChannel(path);
        m_opc = openPackage(m_channel);
        m_streamedParser = createStreamedParser(m_opc);
        return m_streamedParser;
    }

    /**
     * Tries to open an OPCPackage backed by the given channel.
     *
     * @param channel channel to read contents from
     * @return opened {@link OPCPackage} backed by the given channel
     * @throws IOException
     */
    private OPCPackage openPackage(final SeekableByteChannel channel) throws IOException {
        try {
            final var fileType = ExcelUtils.sniffFileType(channel);
            if (fileType == FileMagic.OLE2) {
                // OLE2 file could contain:
                // - encrypted XML content (OOXML)
                // - XLS file with XLSX extension
                final var asm = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
                final var authType = asm.getAuthenticationType();
                if (authType == AuthenticationType.NONE) {
                    channel.close();
                    // This exception would previously be thrown by OPCPackage.open(File),
                    // but is not by OPCPackage.open(InputStream), so we have to throw it ourselves!
                    // The ExcelTableReader catches this and switches to the XLSRead
                    throw new OLE2NotOfficeXmlFileException("OLE2 file but trying to open as OOXML-based file.");
                }
                final var password = asm.getPassword(m_config.getReaderSpecificConfig().getCredentialsProvider());
                return openDecrypting(channel, password, () -> createPasswordIncorrectException(null));
            } else if (fileType == FileMagic.OOXML) {
                // this case means we actually have an unencrypted file, so just open the package without decryption
                return openFromZIP(channel);
            }
            // will be caught and output with a user-friendly error message
            throw new NotOfficeXmlFileException("");
        } catch (GeneralSecurityException | InvalidFormatException e) {
            throw new IOException(e.getMessage(), e);
        }
    }

    @SuppressWarnings("resource") // FSZipPackage will be closed by OPCPackage's close method
    private static OPCPackage openFromZIP(final SeekableByteChannel chan) throws IOException, InvalidFormatException {
        return OPCPackage.open(new FSZipPackage(chan));
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

    /**
     * Creates and returns a {@link AbstractStreamedParserRunnable}.
     *
     * @param opc the {@link OPCPackage} representing the Excel file being read
     * @return the created {@link AbstractStreamedParserRunnable}
     * @throws IOException if an I/O error occurs
     */
    protected abstract AbstractStreamedParserRunnable createStreamedParser(final OPCPackage opc) throws IOException;

    /**
     * @param sheetsData the sheet iterator
     * @param selectedSheet the name of the selected sheet
     * @return the counting input stream of the selected sheet
     * @throws IOException if an I/O error occurs
     * @throws IllegalStateException if the sheet iterator does not contain a sheet with the selected name
     */
    protected static CountingInputStream getSheetStreamWithSheetName(final SheetIterator sheetsData,
        final String selectedSheet) throws IOException {
        while (sheetsData.hasNext()) {
            @SuppressWarnings("resource") // manually closed
            final InputStream sheetIterator = sheetsData.next();
            try {
                if (sheetsData.getSheetName().equals(selectedSheet)) {
                    return new CountingInputStream(sheetIterator);
                } else {
                    sheetIterator.close();
                }
            } catch (Exception e) { // NOSONAR catch anything to make sure the stream is closed
                sheetIterator.close();
            }
        }
        // should never happen as we checked for it already
        throw new IllegalStateException("No sheet with name '" + selectedSheet + "' found.");
    }

    @Override
    public Map<String, Boolean> getSheetNames() {
        return m_sheetNames;
    }

    @Override
    public Set<Integer> getHiddenColumns() {
        return m_streamedParser.getHiddenColumns();
    }

    @Override
    public void closeResources() throws IOException {
        if (m_countingSheetStream != null) {
            m_countingSheetStream.close();
        }
        if (m_opc != null) {
            // read-only OPCPackages should be reverted, not closed, but in case it was opened from an InputStream
            // it is in READ-WRITE mode... (see OPCPackage.open(InputStream))
            // OPCPackage.close checks for that but issues a warning if we're closing a READ-only package
            if (m_opc.getPackageAccess() == PackageAccess.READ) {
                m_opc.revert();
            } else {
                m_opc.close();
            }
        }
        if (m_channel != null) {
            m_channel.close();
            m_channel = null;
        }
    }

    @Override
    public OptionalLong getMaxProgress() {
        return m_sheetSize < 0 ? OptionalLong.empty() : OptionalLong.of(m_sheetSize);
    }

    @Override
    public long getProgress() {
        return m_countingSheetStream.getByteCount();
    }

}
