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
 *   Nov 16, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.table;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.util.FileUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;

/**
 * Class allowing to create and save {@link Workbook}s and create the {@link ExcelTableWriter}. It must be ensured that
 * the {@link ExcelTableWriter} fits with the created {@link Workbook}.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public abstract class WorkbookHandler implements AutoCloseable {

    private static final NodeLogger LOG = NodeLogger.getLogger(WorkbookHandler.class);

    /**
     * The excel format
     */
    protected final ExcelFormat m_format;

    /**
     * The path to the opened excel file.
     */
    protected final Path m_inputPath;

    private Workbook m_workbook;

    private String m_password;

    /**
     * @param format The excel format.
     * @param inputPath The path to an existing excel file which should be loaded into the {@link Workbook}.
     * @param secretPassword password to use for encrypted Excel file, or {@code null} for unencrypted file
     */
    protected WorkbookHandler(final ExcelFormat format, final Path inputPath, final String secretPassword) {
        m_format = format;
        m_inputPath = inputPath;
        m_password = secretPassword;
    }

    /**
     * @return The workbook.
     * @throws IOException
     */
    public Workbook getWorkbook() throws IOException {
        if (m_workbook == null) {
            m_workbook = createWorkbook(m_password);
        }
        return m_workbook;
    }

    /**
     * Creates the {@link Workbook}.
     * @param secretPassword
     *
     * @return the workbook
     * @throws IOException - If the workbook could not be created
     */
    protected abstract Workbook createWorkbook(String secretPassword) throws IOException;

    /**
     * Creates the {@link ExcelTableWriter} associated with the created workbook.
     *
     * @param cfg the {@link ExcelTableConfig} config
     * @param cellWriterFactory the {@link ExcelCellWriterFactory}
     * @return an instance of {@link ExcelTableWriter} that fits with the {@link Workbook} created by
     *         {@link #createWorkbook(String)}
     */
    public ExcelTableWriter createTableWriter(final ExcelTableConfig cfg,
        final ExcelCellWriterFactory cellWriterFactory) {
        CheckUtils.checkState(m_format != null, "Cannot create a table writer before creating a workbook");
        return m_format.createWriter(cfg, cellWriterFactory);
    }

    /**
     * Saves the current {@link Workbook} to the provided file and closes the workbook.
     *
     * @param outPath The path of the target file.
     * @throws IOException
     */
    public void saveFile(final Path outPath) throws IOException {
        var savePath = outPath;

        // if we loaded the Workbook from an existing file, then the Workbook may still hold an
        // open input stream for that file. On some file systems (e.g. SMB) this will lock the
        // file so that it cannot just be overwritten. Hence we need to write the workbook to a temp
        // file first and then close the workbook, to release the lock on the existing file. Then we
        // can replace the original file with the temp file.
        if (m_inputPath != null) {
            savePath = FileUtil.createTempFile("knime-excel-", "").toPath();
        }

        try {
            doSaveFile(savePath);
            close();
            if (savePath != outPath) {
                // this is effectively a file copy, but we can't use Files.copy() because it tries
                // to delete the target file, which fails for the Custom/KNIME URL file system
                try (var in = Files.newInputStream(savePath); var out = Files.newOutputStream(outPath)) {
                    IOUtils.copyLarge(in, out);
                }
            }
        } finally {
            if (savePath != outPath) {
                FSFiles.deleteSafely(savePath);
            }
        }
    }

    private void doSaveFile(final Path path) throws IOException {
        if (m_password == null) {
            writeWorkbook(path);
            return;
        }

        // password present -> write with encryption
        if (m_format == ExcelFormat.XLS) {
            Biff8EncryptionKey.setCurrentUserPassword(m_password);
            writeWorkbook(path);
            Biff8EncryptionKey.setCurrentUserPassword(null);
            return;
        }
        if (m_format == ExcelFormat.XLSX) {
            writeEncryptedWorkbookToXLSX(path);
            return;
        }
        throw new IllegalStateException("Unsupported format: \"%s\"".formatted(m_format));
    }

    private void writeEncryptedWorkbookToXLSX(final Path path) throws IOException {
        // see https://poi.apache.org/encryption.html
        var info = new EncryptionInfo(EncryptionMode.agile);
        var enc = info.getEncryptor();
        enc.confirmPassword(m_password);
        try (final var fs = new POIFSFileSystem()) {
            try (final var os = enc.getDataStream(fs)) {
                m_workbook.write(os);
            } catch (final GeneralSecurityException e) {
                throw new IOException("Encryption of Excel file failed.", e);
            }
            // must close enc output stream before writing file system
            try (final var fos = Files.newOutputStream(path)) {
                fs.writeFilesystem(fos);
            }
        }
    }

    /**
     * Write the workbook to the given path, using the workbook's write method.
     *
     * @param path destination
     * @throws IOException if writing the workbook fails
     */
    private void writeWorkbook(final Path path) throws IOException {
        try (final var out = FSFiles.newOutputStream(path, FileOverwritePolicy.OVERWRITE.getOpenOptions());
                final var buffer = new BufferedOutputStream(out)) {
            m_workbook.write(buffer);
        }
    }

    @Override
    public void close() {
        try {
            if (m_workbook != null) {
                if (m_workbook instanceof SXSSFWorkbook) {
                    ((SXSSFWorkbook)m_workbook).dispose();
                }
                m_workbook.close();
            }
        } catch (IOException e) {
            LOG.debug("Exception while closing Excel workbook: " + e.getMessage(), e);
        } finally {
            m_workbook = null;
        }
    }
}
