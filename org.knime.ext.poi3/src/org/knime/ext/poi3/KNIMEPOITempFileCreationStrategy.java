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
 *   Feb 19, 2018 (ferry): created
 */
package org.knime.ext.poi3;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.util.DefaultTempFileCreationStrategy;
import org.apache.poi.util.TempFileCreationStrategy;
import org.knime.core.util.PathUtils;

/**
 * Temp file creation strategy that ensures that the temp directory always exists (which is a problem in
 * {@link DefaultTempFileCreationStrategy}).
 *
 * @author Ferry Abt, KNIME GmbH, Konstanz, Germany
 */
final class KNIMEPOITempFileCreationStrategy implements TempFileCreationStrategy {
    private final Path m_dir;

    /**
     * Creates the strategy allowing to set the
     *
     * @param dir The directory where the temporary files will be created
     */
    KNIMEPOITempFileCreationStrategy(final Path dir) {
        if (dir == null) {
            throw new IllegalArgumentException("Apache POI temp directory must not be null!");
        }
        m_dir = dir;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public File createTempFile(final String prefix, final String suffix) throws IOException {
        Files.createDirectories(m_dir);

        Path newFile;
        // Set the delete on exit flag, unless explicitly disabled
        if (System.getProperty("poi.keep.tmp.files") == null) {
            newFile = PathUtils.createTempFile(m_dir, prefix, suffix);
        } else {
            newFile = Files.createTempFile(m_dir, prefix, suffix);
        }

        return newFile.toFile();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public File createTempDirectory(final String prefix) throws IOException {
        Files.createDirectories(m_dir);

        Path newDirectory;
        // Set the delete on exit flag, unless explicitly disabled
        if (System.getProperty("poi.keep.tmp.files") == null) {
            newDirectory = PathUtils.createTempDir(prefix, m_dir);
        } else {
            newDirectory = Files.createTempDirectory(m_dir, prefix);
        }

        return newDirectory.toFile();
    }
}
