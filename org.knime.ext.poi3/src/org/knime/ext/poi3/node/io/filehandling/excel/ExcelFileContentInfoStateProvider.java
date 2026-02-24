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
 *   Feb 24, 2026 (Thomas Reifenberger, TNG Technology Consulting GmbH): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.nio.file.Files;
import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.function.Supplier;

import org.apache.poi.EncryptedDocumentException;
import org.knime.base.node.io.filehandling.webui.FileChooserPathAccessor;
import org.knime.base.node.io.filehandling.webui.FileSystemPortConnectionUtil;
import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.updates.StateProvider;

/**
 * Abstract base state provider that reads content information (sheet names, encryption state) from a selected Excel
 * file. Subclasses provide the file selection and encryption types via the generic parameters {@code F} and {@code E},
 * and implement the two abstract methods to create a {@link FileChooserPathAccessor} and resolve the password.
 *
 * @param <F> the type of the file selection
 * @param <E> the type of the encryption settings
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 * @noextend non-public API
 */
@SuppressWarnings("restriction")
public abstract class ExcelFileContentInfoStateProvider<F, E>
    implements StateProvider<ExcelFileContentInfoStateProvider.ExcelFileInfo> {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelFileContentInfoStateProvider.class);

    /** Supplier for the file selection, to be initialised in {@link #initFileAndEncryptionSuppliers}. */
    protected Supplier<F> m_fileSelection;

    /** Supplier for the encryption settings, to be initialised in {@link #initFileAndEncryptionSuppliers}. */
    protected Supplier<E> m_encryption;

    /**
     * Called during {@link #init} after {@code computeAfterOpenDialog()} has been set. Subclasses must assign
     * {@link #m_fileSelection} and {@link #m_encryption} using the provided initializer.
     *
     * @param initializer the state provider initializer
     */
    protected abstract void initFileAndEncryptionSuppliers(StateProviderInitializer initializer);

    /**
     * Creates a {@link FileChooserPathAccessor} for the given file selection and optional file system connection.
     *
     * @param fileSelection the current file selection value
     * @param fsConnection the optional file system port connection
     * @return a new {@link FileChooserPathAccessor}
     */
    @SuppressWarnings("java:S3553") // Optional parameter is part of the established file-handling API
    protected abstract FileChooserPathAccessor createPathAccessor(F fileSelection, Optional<FSConnection> fsConnection);

    /**
     * Resolves the password from the given encryption settings and node parameters input.
     *
     * @param encryption the current encryption settings value
     * @param context the node parameters input (may contain credentials provider)
     * @return the password string, or {@code null} if no password is set
     */
    protected abstract String getPassword(E encryption, NodeParametersInput context);

    @Override
    public final void init(final StateProviderInitializer initializer) {
        /*
         * Looking up sheet names can take time, especially for large files or remote file systems
         * -> do it after opening the dialog (otherwise the dialog opening is blocked)
         */
        initializer.computeAfterOpenDialog();
        initFileAndEncryptionSuppliers(initializer);
    }

    @Override
    public final ExcelFileInfo computeState(final NodeParametersInput context) {
        final var fsConnection = FileSystemPortConnectionUtil.getFileSystemConnection(context);
        try (final var accessor = createPathAccessor(m_fileSelection.get(), fsConnection)) {
            final var path = accessor.getPaths(s -> {
            }).get(0);
            if (!Files.exists(path)) {
                return new ExcelFileInfo(false, Collections.emptyList());
            }
            final String password = getPassword(m_encryption.get(), context);
            try (final var inputStream = FSFiles.newInputStream(path)) {
                final var sheetNames = ExcelUtils.readSheetNames(inputStream, password);
                return new ExcelFileInfo(false, sheetNames);
            }
        } catch (final EncryptedDocumentException e) { // NOSONAR swallowing exception is intentional
            return new ExcelFileInfo(true, Collections.emptyList());
        } catch (final Exception e) {
            LOGGER.warn("Could not read sheet names for suggestions.", e);
            return new ExcelFileInfo(false, Collections.emptyList());
        }
    }

    /**
     * Holds the information read from the Excel input file: sheet names and whether a password is missing.
     */
    public static class ExcelFileInfo {

        private final List<String> m_sheetNames;

        private final boolean m_isPasswordMissing;

        /**
         * Creates a new {@link ExcelFileInfo} with the given password state and sheet names.
         *
         * @param isPasswordMissing {@code true} if the file is encrypted and no valid password was supplied
         * @param sheetNames the list of sheet names read from the file
         */
        public ExcelFileInfo(final boolean isPasswordMissing, final List<String> sheetNames) {
            m_isPasswordMissing = isPasswordMissing;
            m_sheetNames = sheetNames;
        }

        /**
         * @return {@code true} if the file is encrypted and no valid password was supplied
         */
        public boolean isPasswordMissing() {
            return m_isPasswordMissing;
        }

        /**
         * @return the sheet names read from the file, or an empty list if the file could not be read
         */
        public List<String> getSheetNames() {
            return m_sheetNames;
        }
    }
}
