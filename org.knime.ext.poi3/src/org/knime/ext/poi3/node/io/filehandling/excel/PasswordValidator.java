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
 *   25 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import java.util.function.BiFunction;
import java.util.function.Supplier;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.defaultnodesettings.filechooser.StatusMessageReporter;
import org.knime.filehandling.core.defaultnodesettings.status.DefaultStatusMessage;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessageUtils;

/**
 * StatusMessageReporter that checks that the file at the given path can be decrypted using the given password.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class PasswordValidator implements StatusMessageReporter {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(PasswordValidator.class);

    private final Path m_path;

    private final Supplier<String> m_passwordSupplier;

    private final Supplier<ExcelFormat> m_expectedFormat;

    private final BiFunction<Path, EncryptedDocumentException, StatusMessage> m_decryptionErrorHandler;

    /**
     * Constructor.
     *
     * @param path path of file to check
     * @param passwordSupplier supplier for password to check
     * @param expectedFormat expected excel format
     * @param decryptionErrorHandler handler to use when the decryption fails (wrong/no password)
     */
    public PasswordValidator(final Path path, final Supplier<String> passwordSupplier,
        final Supplier<ExcelFormat> expectedFormat,
        final BiFunction<Path, EncryptedDocumentException, StatusMessage> decryptionErrorHandler) {
        m_path = path;
        m_passwordSupplier = passwordSupplier;
        m_expectedFormat = expectedFormat;
        m_decryptionErrorHandler = decryptionErrorHandler;
    }

    @Override
    public StatusMessage report() throws IOException, InvalidSettingsException {
        LOGGER.debug("Refreshing file decryption status...");
        try (final var in = FSFiles.newInputStream(m_path); final var buf = FileMagic.prepareToCheckMagic(in)) {
            final var fileType = FileMagic.valueOf(buf);
            if (fileType != FileMagic.OOXML && fileType != FileMagic.OLE2) {
                return DefaultStatusMessage.mkError("Unsupported file type \"%s\"", fileType);
            }
            final var expected = m_expectedFormat.get();
            if (fileType == FileMagic.OOXML) {
                // cannot be encrypted, otherwise would be OLE2
                return expectFormat(expected, ExcelFormat.XLSX).orElse(StatusMessageUtils.SUCCESS_MSG);
            }
            // encrypted XLSX or XLS
            try (final var fs = new POIFSFileSystem(buf)) {
                final var root = fs.getRoot();
                final var fmtError =
                    expectFormat(expected, CryptUtil.isEncryptedOOXML(root) ? ExcelFormat.XLSX : ExcelFormat.XLS);
                if (fmtError.isPresent()) { // NOSONAR cannot use orElseGet since supplier cannot throw our checked exception...
                    return fmtError.get();
                }
                // verify password only after the file type is what we expects and could need a password
                return verifyPassword(m_path, fs);
            }
        }
    }

    private StatusMessage verifyPassword(final Path path, final POIFSFileSystem fs) throws IOException {
        try {
            CryptUtil.verifyPasswordForOLE2(fs, m_passwordSupplier.get());
            return StatusMessageUtils.SUCCESS_MSG;
        } catch (EncryptedDocumentException e) {
            return m_decryptionErrorHandler.apply(path, e);
        }
    }

    private static Optional<StatusMessage> expectFormat(final ExcelFormat expected, final ExcelFormat actual) {
        if (expected == null || expected == actual) {
            return Optional.empty();
        }
        return Optional.ofNullable(DefaultStatusMessage.mkError("Unable to append to existing file: "
            + "you have selected \"%s\" but the existing file is actually of type \"%s\"", expected, actual));
    }

}
