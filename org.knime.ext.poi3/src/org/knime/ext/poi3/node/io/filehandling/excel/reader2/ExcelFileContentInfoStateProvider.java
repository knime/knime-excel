package org.knime.ext.poi3.node.io.filehandling.excel.reader2;

import org.apache.poi.EncryptedDocumentException;
import org.knime.base.node.io.filehandling.webui.FileChooserPathAccessor;
import org.knime.base.node.io.filehandling.webui.FileSystemPortConnectionUtil;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionParameters;
import org.knime.core.node.NodeLogger;
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
            final var password = m_encryption.get().m_credentials.getPassword();
            if (!Files.exists(path)) {
                return new ExcelFileInfo(false, Collections.emptyList());
            }
            try (final var inputStream = FSFiles.newInputStream(path)) {
                CryptUtil.verifyPassword(inputStream, password);
            }
            try (final var inputStream = FSFiles.newInputStream(path)) {
                final var sheetNames = ExcelUtils.readSheetNames(inputStream, password);
                return new ExcelFileInfo(false, sheetNames);
            }
        } catch (final EncryptedDocumentException e) {
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
