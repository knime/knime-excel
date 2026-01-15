package org.knime.ext.poi3.node.io.filehandling.excel.reader2;

import java.util.Optional;
import java.util.function.Supplier;

import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.StateProvider;
import org.knime.node.parameters.widget.message.TextMessage;

@Layout(ExcelFileInfoMessage.FileInfoMessageLayout.class)
class ExcelFileInfoMessage implements NodeParameters {

    @After(ExcelTableReaderParameters.FileAndSheetSection.Source.class)
    interface FileInfoMessageLayout {
    }

    @TextMessage(FileInfoMessageProvider.class)
    Void m_fileInfoMessage;

    private static class FileInfoMessageProvider implements StateProvider<Optional<TextMessage.Message>> {

        Supplier<ExcelFileContentInfoStateProvider.ExcelFileInfo> m_excelFileInfo;

        @Override
        public void init(final StateProviderInitializer initializer) {
            m_excelFileInfo = initializer.computeFromProvidedState(ExcelFileContentInfoStateProvider.class);
        }

        @Override
        public Optional<TextMessage.Message> computeState(NodeParametersInput parametersInput) {
            if (m_excelFileInfo.get().isPasswordMissing()) {
                return Optional.of(new TextMessage.Message("Password required",
                    "Found an encrypted file that cannot be opened with the provided credentials "
                        + "(see advanced settings).",
                    TextMessage.MessageType.INFO));
            }
            return Optional.empty();
        }
    }
}
