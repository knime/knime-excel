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

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.GeneralSecurityException;
import java.util.Map;
import java.util.OptionalLong;
import java.util.Set;
import java.util.function.Consumer;

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
import org.knime.core.util.FileUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
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
     * @param sheetNamesConsumer
     * @throws IOException if an I/O exception occurs
     */
    protected AbstractStreamedRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
            final Consumer<Map<String, Boolean>> sheetNamesConsumer)
        throws IOException {
        super(path, config, sheetNamesConsumer);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    /** The size of the sheet to read. */
    protected long m_sheetSize;

    private AbstractStreamedParserRunnable m_streamedParser;


    @Override
    protected ExcelParserRunnable createParser(final File file) throws IOException {
        try {
            m_streamedParser = switch (FileMagic.valueOf(file)) {
                case OLE2 -> createParserFromOLE2(file);
                case OOXML -> createParserFromOOXML(file);
                // will be caught and output with a user-friendly error message
                default -> throw new NotOfficeXmlFileException("");
            };
        } catch (GeneralSecurityException | InvalidFormatException e) {
            throw new IOException(e.getMessage(), e);
        }
        return m_streamedParser;
    }

    @SuppressWarnings("resource") // pkg ownership handed to parser
    private AbstractStreamedParserRunnable createParserFromOOXML(final File file)
            throws InvalidFormatException, IOException {
        final var pkg = OPCPackage.open(file, PackageAccess.READ);
        try {
            // ownership of pkg handed to parser
            return createStreamedParser(pkg);
        } catch (final IOException e) {
            // if there is a problem creating the parser, we need to close the package
            pkg.revert();
            throw e;
        }
    }

    private AbstractStreamedParserRunnable createParserFromOLE2(final File file)
            throws IOException, GeneralSecurityException, InvalidFormatException {
        // encrypted OOXML files are stored as encrypted OLE2 files that contain the xml content
        // so reaching this point, this XLSX file is encrypted (or it is not a valid XLSX file in the first place)
        final var authModel = m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
        try (final var fs = new POIFSFileSystem(file)) {
            if (!CryptUtil.isEncryptedOOXML(fs.getRoot())) {
                throw new OLE2NotOfficeXmlFileException("The given file is not an encrypted XLSX file.");
            }
            final var password =
                    CryptUtil.getPassword(authModel, m_config.getReaderSpecificConfig().getCredentialsProvider());
            if (password == null) {
                throw createPasswordProtectedFileException(null);
            }
            final var info = new EncryptionInfo(fs);
            final var d = Decryptor.getInstance(info);
            if (!d.verifyPassword(password)) {
                throw createPasswordIncorrectException(null);
            }
            try (final var decryptedStream = d.getDataStream(fs)) {
                // We cache the contents to a temporary file, since otherwise encrypted Excel files would be buffered in
                // memory fully, since we'd have to use OPCPackage.open(InputStream).
                // Using `AesZipFileZipEntrySource.createZipEntrySource(decryptedStream)` to buffer the file on disk
                // encrypted fails with "Truncated ZIP file"
                final var tempFile = FileUtil.createTempFile("tempXlsx", ".xlsx", FileUtil.getWorkflowTempDir(), true);
                try {
                    Files.copy(decryptedStream, tempFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    return createParserFromOOXML(tempFile);
                } catch (final IOException | InvalidFormatException e) {
                    Files.deleteIfExists(tempFile.toPath());
                    throw e;
                }
            }
        }
    }

    /**
     * Creates and returns a {@link AbstractStreamedParserRunnable}.
     *
     * @param pkg package representing the Excel 2007+ file to parse
     * @return the created {@link AbstractStreamedParserRunnable}
     * @throws IOException if an I/O error occurs
     */
    protected abstract AbstractStreamedParserRunnable createStreamedParser(final OPCPackage pkg) throws IOException;

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
    public Set<Integer> getHiddenColumns() {
        return m_streamedParser.getHiddenColumns();
    }

    @Override
    public OptionalLong getMaxProgress() {
        return m_sheetSize < 0 ? OptionalLong.empty() : OptionalLong.of(m_sheetSize);
    }

    @Override
    public long getProgress() {
        return m_streamedParser.getProgress();
    }

}
