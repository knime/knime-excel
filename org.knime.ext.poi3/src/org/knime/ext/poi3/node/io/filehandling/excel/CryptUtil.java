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
 *   16 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.StringJoiner;
import java.util.stream.Collectors;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.record.RecordFactoryInputStream;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.node.workflow.CredentialsProvider;

/**
 * Utility for dealing with password protected Excel files.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class CryptUtil {

    private static final String PASSWD_INCORRECT = "The given password is incorrect";

    private static final String PASSWD_MISSING = "The password is missing";

    private CryptUtil() {
        // hidden
    }

    /**
     * Verifies the given password against the Excel file in the given input.
     * Consider using the specialized method {@link #verifyPassword(FileMagic, InputStream, String)} if you already
     * know the file type or {@code .verifyPasswordFor*(...)} if you already have open resources.
     *
     * @param in input stream containing the Excel file
     * @param password password to verify
     * @throws EncryptedDocumentException
     * @throws IOException
     */
    public static void verifyPassword(final InputStream in, final String password)
            throws EncryptedDocumentException, IOException {
        try (final var buf = FileMagic.prepareToCheckMagic(in)) {
            final var fileType = FileMagic.valueOf(buf);
            verifyPassword(fileType, buf, password);
        }
    }

    /**
     * Verify the given password against the given input, expecting the input to correspond to the given file type,
     * which can be either {@link FileMagic#OOXML} or {@link FileMagic#OLE2}.
     *
     * @param fileType
     * @param in
     * @param password
     * @throws EncryptedDocumentException
     * @throws IOException
     */
    public static void verifyPassword(final FileMagic fileType, final InputStream in, final String password)
            throws EncryptedDocumentException, IOException {
        if (fileType == FileMagic.OOXML) {
            // unencrypted XLSX file
            return;
        }

        if (fileType == FileMagic.OLE2) {
            try (final var fs = new POIFSFileSystem(in)) {
                verifyPasswordForOLE2(fs, password);
                return;
            }
        }
        throw new IllegalArgumentException("No support for file of type: " + fileType);
    }

    /**
     * Verify the given password against the workbook contained in the given POI file system.
     * @param fs file system to look in for workbook
     * @param password password to verify
     * @throws IOException if decryption has a problem
     * @throws EncryptedDocumentException if the password is missing or incorrect
     */
    @SuppressWarnings("unused")
    public static void verifyPasswordForOLE2(final POIFSFileSystem fs, final String password)
        throws EncryptedDocumentException, IOException {
        final var root = fs.getRoot();
        if (isEncryptedOOXML(root)) {
            verifyPasswordForEncryptedOOXML(root, password);
            return;
        }
        // (encrypted) XLS?
        try (final var docIn = findDocumentInput(root)) {
            Biff8EncryptionKey.setCurrentUserPassword(password);
            // This next line checks for encryption and automatically opens a decrypted stream,
            // which fails if the password is missing or incorrect. This seems like the most resource-efficient
            // way to check if a .xls is password protected.
            new RecordFactoryInputStream(docIn, false);
            Biff8EncryptionKey.setCurrentUserPassword(null);
        } catch (final EncryptedDocumentException e) {
            // password or default password invalid
            throw new EncryptedDocumentException(password != null ? PASSWD_INCORRECT : PASSWD_MISSING, e);
        }
    }

    /**
     * Verifies that the given password can be used to decrypt the workbook contents (incl. using the default password
     * when no password is passed) and returns a decryptor.
     *
     * @param root directory root node
     * @param password password to decrypt with
     * @return decryptor that can be used to decrypt the content
     *
     * @throws EncryptedDocumentException if the password is incorrect or missing
     * @throws IOException if password verification or decryption throws an error
     */
    public static Decryptor verifyPasswordForEncryptedOOXML(final DirectoryNode root, final String password)
            throws EncryptedDocumentException, IOException {
        final var info = new EncryptionInfo(root);
        final var decryptor = Decryptor.getInstance(info);
        try {
            if (password != null) {
                if (decryptor.verifyPassword(password)) {
                    return decryptor;
                }
                throw new EncryptedDocumentException(PASSWD_INCORRECT);
            }
            // password is null, so check with the default password
            if (decryptor.verifyPassword(Decryptor.DEFAULT_PASSWORD)) {
                return decryptor;
            }
            throw new EncryptedDocumentException(PASSWD_MISSING);
        } catch (final GeneralSecurityException e) {
            throw new IOException(e);
        }
    }

    /**
     * Find the workbook document in the given directory node (should be the root of the POIFS).
     *
     * @param root directory to search for Workbook entry and open stream from
     * @return document input stream if Workbook entry could be found
     * @throws IOException if creating the document stream fails
     * @throws FileNotFoundException if the directory node did not contain a workbook entry
     */
    private static DocumentInputStream findDocumentInput(final DirectoryNode root) throws IOException {
        final var candidates = InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES;
        for (final var c : candidates) {
            if (root.hasEntry(c)) {
                return root.createDocumentInputStream(root.getEntry(c));
            }
        }
        // none of the entries we expect are contained in the POIFS...
        final var sb = new StringBuilder("The input did not contain a Workbook, had: ");
        final var sj = new StringJoiner(", ", "[", "]");
        final var entries = root.getEntryNames();
        for (final var e : entries) {
            sj.add(e);
        }
        sb.append(sj.toString());
        throw new FileNotFoundException(sb.toString());
    }

    /**
     * Checks if the given directory node contains an encrypted OOXML document.
     *
     * @param root directory node to check
     * @return {@code true} if the directory node contains an encrypted OOXML document, {@code false} otherwise
     */
    public static boolean isEncryptedOOXML(final DirectoryEntry root) {
        return root.hasEntry(Decryptor.DEFAULT_POIFS_ENTRY);
    }

    /**
     * Gets the decrypted OOXML stream from the directory root or {@code null} if the document is not an encrypted OOXML
     * document.
     *
     * @param root root directory node
     * @param password passoword used for decryption
     * @return input stream containing OOXML document, or {@code null} if not encrypted
     * @throws IOException if decryption has a problem
     * @throws EncryptedDocumentException if the password is missing or invalid
     */
    public static InputStream getDecryptedStream(final DirectoryNode root, final String password) throws IOException {
        if (!isEncryptedOOXML(root)) {
            return null;
        }
        final var decryptor = verifyPasswordForEncryptedOOXML(root, password);
        try {
            return decryptor.getDataStream(root);
        } catch (GeneralSecurityException e) {
            throw new IOException(e);
        }
    }

    /**
     * Gets the password used for encryption/decryption of Excel files based on the given authentication settings model
     * and credential provider.
     * In case the authentication model uses credentials but the credentials provider is {@code null}, the returned
     * password will also be {@code null}.
     *
     * @param authModel non-{@code null} authentication settings model
     * @param cp nullable credentials provider
     * @return the password or {@code null} if no password was supplied via the authentication model and credentials
     *          provider
     */
    public static String getPassword(final SettingsModelAuthentication authModel, final CredentialsProvider cp) {
        final var authType = authModel.getAuthenticationType();
        CheckUtils.check(switch(authType) {
            case NONE, CREDENTIALS, PWD -> true;
            case KERBEROS, USER, USER_PWD -> false;
        }, IllegalArgumentException::new, () -> {
            final var sup = List.of(AuthenticationType.NONE, AuthenticationType.CREDENTIALS, AuthenticationType.PWD)
                    .stream().map(AuthenticationType::getText).collect(Collectors.joining(", "));
            return String.format("Unsupported authentication type \"%s\", supported: %s", authType.getText(), sup);
        });

        return authType == AuthenticationType.NONE
            // avoids null pointer exception when getting password without credentials provider
            || (authType == AuthenticationType.CREDENTIALS && cp == null) ? null : authModel.getPassword(cp);
    }
}
