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
 *   14 Aug 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.io.IOException;
import java.io.InputStream;
import java.nio.channels.SeekableByteChannel;
import java.util.Enumeration;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipEntrySource;

/**
 * An adapter to open a {@link SeekableByteChannel}, for example one obtained
 * through {@code Files.newByteChannel(fsPath)}, with {@link OPCPackage#open(ZipEntrySource)}.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class FSZipPackage implements ZipEntrySource {

    private ZipFile m_zip;

    /**
     * Opens the given channel for reading, assuming "UTF8" for file names, like
     * {@link ZipFile#ZipFile(SeekableByteChannel)}.
     *
     * @param chann channel to read ZIP entries from
     * @throws IOException if an error occurs when opening the channel
     * @see ZipFile#ZipFile(SeekableByteChannel)
     */
    public FSZipPackage(final SeekableByteChannel chann) throws IOException {
        m_zip = new ZipFile(chann);
    }

    @Override
    public Enumeration<? extends ZipArchiveEntry> getEntries() {
        return m_zip.getEntries();
    }

    @Override
    public ZipArchiveEntry getEntry(final String path) {
        return m_zip.getEntry(path);
    }

    @Override
    public InputStream getInputStream(final ZipArchiveEntry entry) throws IOException {
        return m_zip.getInputStream(entry);
    }

    @Override
    public void close() throws IOException {
        if (!isClosed()) {
            // this also closes the channel
            m_zip.close();
            m_zip = null;
        }
    }

    @Override
    public boolean isClosed() {
        return m_zip == null;
    }

}
