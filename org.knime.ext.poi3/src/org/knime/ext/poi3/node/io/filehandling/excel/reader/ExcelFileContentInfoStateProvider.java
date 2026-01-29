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
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.apache.poi.EncryptedDocumentException;
import org.knime.base.node.io.filehandling.webui.FileChooserPathAccessor;
import org.knime.base.node.io.filehandling.webui.FileSystemPortConnectionUtil;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionParameters;
import org.knime.core.node.NodeLogger;
import org.knime.core.webui.node.dialog.defaultdialog.NodeParametersInputImpl;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.DefaultFileChooserFilters;
import org.knime.core.webui.node.dialog.defaultdialog.internal.file.MultiFileSelection;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.updates.StateProvider;

import java.nio.file.Files;
import java.util.Collections;
import java.util.List;
import java.util.function.Supplier;

/**
 * Intermediate state provider reading all relevant information from the selected Excel file (or the first one for a
 * multi-file selection). The result is used by other state providers to provide choices etc. in the UI.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
class ExcelFileContentInfoStateProvider implements StateProvider<ExcelFileContentInfoStateProvider.ExcelFileInfo> {
    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelFileContentInfoStateProvider.class);

    Supplier<MultiFileSelection<DefaultFileChooserFilters>> m_fileSelection;

    Supplier<EncryptionParameters> m_encryption;

    @Override
    public void init(final StateProvider.StateProviderInitializer initializer) {
        /*
         * Looking up sheet names can take time, especially for large files or remote file systems
         * -> do it after opening the dialog (otherwise the dialog opening is blocked)
         */
        initializer.computeAfterOpenDialog();
        m_fileSelection = initializer.computeFromValueSupplier(MultiFileSelectionParameters.FileSelectionRef.class);
        m_encryption = initializer.computeFromValueSupplier(ExcelTableReaderParameters.EncryptionParametersRef.class);
    }

    @Override
    public ExcelFileInfo computeState(final NodeParametersInput context) {

        var fsConnection = FileSystemPortConnectionUtil.getFileSystemConnection(context);
        // Note: there is some duplication, as the reader framework also instantiates a FileChooserPathAccessor
        // If this becomes a performance issue, we could think about introducing an intermediate state provider
        try (final var accessor = new FileChooserPathAccessor(m_fileSelection.get(), fsConnection)) {
            final var path = accessor.getPaths(s -> {
            }).get(0);
            if (!Files.exists(path)) {
                return new ExcelFileInfo(false, Collections.emptyList());
            }
            final String password = switch (m_encryption.get().m_encryptionMode) {
                case NO_ENCRYPTION -> null;
                case PASSWORD -> {
                    var credentialsProvider = ((NodeParametersInputImpl)context).getCredentialsProvider().orElseThrow();
                    yield m_encryption.get().m_credentials.toCredentials(credentialsProvider).getPassword();
                }
            };
            try (final var inputStream = FSFiles.newInputStream(path)) {
                CryptUtil.verifyPassword(inputStream, password);
            }
            try (final var inputStream = FSFiles.newInputStream(path)) {
                final var sheetNames = ExcelUtils.readSheetNames(inputStream, password);
                return new ExcelFileInfo(false, sheetNames);
            }
        } catch (final EncryptedDocumentException e) { // NOSONAR swallowing exception is intentional
            return new ExcelFileInfo(true, Collections.emptyList());
        } catch (final Exception e) {
            // in case of any error, return no choices
            LOGGER.warn("Could not read sheet names for suggestions.", e);
            return new ExcelFileInfo(false, Collections.emptyList());
        }
    }

    static class ExcelFileInfo {
        private final List<String> m_sheetNames;

        private final boolean m_isPasswordMissing;

        public ExcelFileInfo(boolean isPasswordMissing, List<String> sheetNames) {
            m_isPasswordMissing = isPasswordMissing;
            m_sheetNames = sheetNames;
        }

        boolean isPasswordMissing() {
            return m_isPasswordMissing;
        }

        List<String> getSheetNames() {
            return m_sheetNames;
        }
    }

}
