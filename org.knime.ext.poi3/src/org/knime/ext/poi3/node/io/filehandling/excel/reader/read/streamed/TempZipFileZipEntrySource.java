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
 *   26 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.streamed;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.Functions.FailableRunnable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipFileZipEntrySource;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.knime.core.node.NodeLogger;
import org.knime.core.util.FileUtil;

/**
 * Temporary zip file source for {@link OPCPackage}, whose underlying temp file is deleted when the source is closed (or
 * the JVM exits).
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class TempZipFileZipEntrySource extends ZipFileZipEntrySource {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(TempZipFileZipEntrySource.class);

    private final FailableRunnable<IOException> m_closeAction;

    private TempZipFileZipEntrySource(final ZipSecureFile zip, final FailableRunnable<IOException> closeAction) {
        super(zip);
        m_closeAction = closeAction;
    }

    /**
     * Creates a new temporary zip entry source from the given input stream (which must contain the contents of a valid
     * ZIP file).
     *
     * @param inputStream contents of ZIP file
     * @return temporary zip entry source
     * @throws IOException
     */
    @SuppressWarnings("resource") // zip file will be closed by zip entry source
    static TempZipFileZipEntrySource create(final InputStream inputStream) throws IOException {
        // We cache the contents to a temporary file, since otherwise encrypted Excel files would be buffered in memory
        // fully, since we'd have to use OPCPackage.open(InputStream).
        // Using `AesZipFileZipEntrySource.createZipEntrySource(decryptedStream)` to buffer the file on disk encrypted
        // fails with "Truncated ZIP file" (but example from poi works...)
        final var tempFile = FileUtil.createTempFile("tempZipSource", ".zip", FileUtil.getWorkflowTempDir(), true);
        LOGGER.debug("Caching password protected file temporarily");
        FileUtils.copyInputStreamToFile(inputStream, tempFile);
        return new TempZipFileZipEntrySource(new ZipSecureFile(tempFile), () -> {
            LOGGER.debug("Cleaning up temporary file");
            if (!Files.deleteIfExists(tempFile.toPath())) {
                LOGGER.warn("Temporary file \"%s\" could not be removed (or was already removed)."
                    .formatted(tempFile.getAbsolutePath()));
            }
        });
    }

    @Override
    public void close() throws IOException {
        final var closed = isClosed();
        super.close();
        if (!closed) {
            m_closeAction.run();
        }
    }

}
