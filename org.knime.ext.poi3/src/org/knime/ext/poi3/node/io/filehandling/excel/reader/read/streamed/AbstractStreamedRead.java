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

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.Map;
import java.util.OptionalLong;
import java.util.Set;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.NamedRange;
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

    /** The opc package containing/holding the sheet that is read. */
    protected OPCPackage m_opc;

    /** The counting sheet stream. */
    protected CountingInputStream m_countingSheetStream;

    /** The size of the sheet to read. */
    protected long m_sheetSize;

    /** The map of sheet names with names as keys being in the same order as in the workbook. */
    protected Map<String, Boolean> m_sheetNames;

    private AbstractStreamedParserRunnable m_streamedParser;

    private POIFSFileSystem m_poiFS;

    protected Map<String, NamedRange> m_namedRanges;

    @SuppressWarnings("resource") // resources closed in #closeResources
    @Override
    protected ExcelParserRunnable createParser(final InputStream inputStream) throws IOException {
        m_streamedParser = createStreamedParser(checkEncryptionAndDecrypt(inputStream));
        return m_streamedParser;
    }

    private InputStream checkEncryptionAndDecrypt(final InputStream inputStream) throws IOException {
        final SettingsModelAuthentication authenticationSettingsModel =
            m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
        final AuthenticationType authenticationType = authenticationSettingsModel.getAuthenticationType();
        if (authenticationType == AuthenticationType.NONE) {
            return inputStream;
        }

        // we need an input stream that supports #mark and #reset
        final BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);
        switch (FileMagic.valueOf(bufferedInputStream)) {
            case OLE2:
                // encrypted OOXML files are stored as encrypted OLE2 files that contain the xml content
                return decrypt(bufferedInputStream);
            case OOXML:
                // this case means we actually have a not encrypted file, so just return the stream without decryption
                return bufferedInputStream;
            default:
                // will be caught and output with a user-friendly error message
                throw new NotOfficeXmlFileException("");
        }
    }

    private InputStream decrypt(final BufferedInputStream bufferedInputStream) throws IOException {
        try {
            m_poiFS = new POIFSFileSystem(bufferedInputStream);
            final EncryptionInfo encryptionInfo = new EncryptionInfo(m_poiFS);
            final Decryptor d = Decryptor.getInstance(encryptionInfo);
            final SettingsModelAuthentication authenticationSettingsModel =
                m_config.getReaderSpecificConfig().getAuthenticationSettingsModel();
            final String password =
                authenticationSettingsModel.getPassword(m_config.getReaderSpecificConfig().getCredentialsProvider());
            if (!d.verifyPassword(password)) {
                bufferedInputStream.close();
                m_poiFS.close();
                throw createPasswordIncorrectException(null);
            }
            return d.getDataStream(m_poiFS);
        } catch (GeneralSecurityException e) {
            throw new IOException(e.getMessage(), e);
        }
    }

    /**
     * Creates and returns a {@link AbstractStreamedParserRunnable}.
     *
     * @param inputStream the input stream of the file being read
     * @return the created {@link AbstractStreamedParserRunnable}
     * @throws IOException if an I/O error occurs
     */
    protected abstract AbstractStreamedParserRunnable createStreamedParser(InputStream inputStream) throws IOException;

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
    public Map<String, NamedRange> getNamedRanges() {
        return m_namedRanges;
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
            m_opc.close();
        }
        if (m_poiFS != null) {
            m_poiFS.close();
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
