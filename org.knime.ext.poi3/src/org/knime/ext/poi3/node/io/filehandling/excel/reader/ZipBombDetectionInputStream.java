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
 *   Jan 5, 2021 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.compress.utils.InputStreamStatistics;
import org.apache.poi.util.IOUtils;

/**
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public class ZipBombDetectionInputStream extends FilterInputStream {

    // don't alert for expanded sizes smaller than 100k
    private static final long GRACE_ENTRY_SIZE = 100 * 1024L;

    private static final double MIN_INFLATE_RATIO = 0.01d;

    private static final long MAX_ENTRY_SIZE = 0xFFFFFFFFL;

    private static final String MAX_ENTRY_SIZE_MSG =
        "Zip bomb detected! The file would exceed the max size of the expanded data in the zip-file.\n"
            + "This may indicates that the file is used to inflate memory usage and thus could pose a security risk.\n"
            + "You can adjust this limit via ZipSecureFile.setMaxEntrySize() if you need to work with files which are very large.\n"
            + "Uncompressed size: %d, Raw/compressed size: %d\n" + "Limits: MAX_ENTRY_SIZE: %d, Entry: %s";

    private static final String MIN_INFLATE_RATIO_MSG =
        "Zip bomb detected! The file would exceed the max. ratio of compressed file size to the size of the expanded data.\n"
            + "This may indicate that the file is used to inflate memory usage and thus could pose a security risk.\n"
            + "You can adjust this limit via ZipSecureFile.setMinInflateRatio() if you need to work with files which exceed this limit.\n"
            + "Uncompressed size: %d, Raw/compressed size: %d, ratio: %f\n"
            + "Limits: MIN_INFLATE_RATIO: %f, Entry: %s";

    public ZipBombDetectionInputStream(final InputStream is) {
        super(new ZipArchiveInputStream(is));
//        if (!(is instanceof InputStreamStatistics)) {
//            throw new IllegalArgumentException(
//                "InputStream of class " + is.getClass() + " is not implementing InputStreamStatistics.");
//        }
    }

    @Override
    public int read() throws IOException {
        int b = super.read();
        if (b > -1) {
            checkThreshold();
        }
        return b;
    }

    @Override
    public int read(final byte[] b, final int off, final int len) throws IOException {
        int cnt = super.read(b, off, len);
        if (cnt > -1) {
            checkThreshold();
        }
        return cnt;
    }

    @Override
    public long skip(final long n) throws IOException {
        long cnt = IOUtils.skipFully(super.in, n);
        if (cnt > 0) {
            checkThreshold();
        }
        return cnt;
    }

    private void checkThreshold() throws IOException {
        //        if (!guardState) {
        //            return;
        //        }

        final InputStreamStatistics stats = (InputStreamStatistics)in;
        final long payloadSize = stats.getUncompressedCount();
        final long rawSize = stats.getCompressedCount();
        final String entryName = "asd"; // entry == null ? "not set" : entry.getName();

        // check the file size first, in case we are working on uncompressed streams
        if (payloadSize > MAX_ENTRY_SIZE) {
            throw new IOException(
                String.format(Locale.ROOT, MAX_ENTRY_SIZE_MSG, payloadSize, rawSize, MAX_ENTRY_SIZE, entryName));
        }

        // don't alert for small expanded size
        if (payloadSize <= GRACE_ENTRY_SIZE) {
            return;
        }

        double min = 0.6;
        double ratio = rawSize / (double)payloadSize;
        if (ratio >= min) {
            return;
        }

        // one of the limits was reached, report it
        throw new IOException(String.format(Locale.ROOT, MIN_INFLATE_RATIO_MSG, payloadSize, rawSize, ratio,
            min, entryName));
    }
}
