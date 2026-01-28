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
 *   14 May 2025 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.File;
import java.io.IOException;
import java.nio.channels.ClosedByInterruptException;
import java.nio.file.Path;
import java.util.Map;
import java.util.OptionalLong;
import java.util.Set;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import org.junit.jupiter.api.Test;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelReaderTestHelper;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;

/**
 * Tests for the {@link ExcelParserRunnable} exception handling. Actual parsing is tested via Testflows.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
class ExcelParserRunnableTest {

    private static final TableReadConfig<ExcelTableReaderConfig> CFG = ExcelReaderTestHelper.createReadConfig();

    private static final Path UNUSED_PATH = Path.of("unused");

    /**
     * Method that tests the propagation of an unhandled exception.
     */
    @Test
    void testHandlingOfUnhandledException() {
        assertThrows(UnhandledException.class, () -> { // NOSONAR test code
            try (final var read = new TestRead(UNUSED_PATH, CFG, ExcelReaderTestHelper.SHEET_NAMES_DUMMY) {

                @Override
                protected void throwException() {
                    throw new UnhandledException();
                }

            }) {
                assertStartEnd(read);
            } // closing of read should evaluate the throwable set by parsing and pass on any unhandled exception
        }, "Expected to propagate unhandled exception");
    }

    /**
     * Method that tests handling of plain {@link ClosedByInterruptException} (case AP-22737).
     *
     * @throws IOException when read has an unexpected problem
     * @throws InterruptedException when our test code is interrupted
     */
    @Test
    void testHandlingOfClosedByInterruptException() throws IOException, InterruptedException {
        try (final var read = new TestRead(UNUSED_PATH, CFG, ExcelReaderTestHelper.SHEET_NAMES_DUMMY) {

            @Override
            protected void throwException() throws Exception {
                // XLSX reading could throw this
                // this exception is supposed to be handled
                throw new ClosedByInterruptException();
            }

        }) {
            assertStartEnd(read);
        } // closing of read should evaluate the throwable set by parsing and pass on any unhandled exception
    }

    /**
     * Method that tests handling of wrapped {@link ClosedByInterruptException} (case AP-24316).
     *
     * @throws IOException when read has an unexpected problem
     * @throws InterruptedException when our test code is interrupted
     */
    @Test
    void testHandlingOfWrappedClosedByInterruptException() throws IOException, InterruptedException {
        try (final var read = new TestRead(UNUSED_PATH, CFG, ExcelReaderTestHelper.SHEET_NAMES_DUMMY) {

            @Override
            protected void throwException() {
                // wrap CBIE exactly like LittleEndianInputStream used by XLSB reading
                // this exception is supposed to be handled
                throw new IllegalStateException(new ClosedByInterruptException());
            }

        }) {
            assertStartEnd(read);
        } // closing of read should evaluate the throwable set by parsing and pass on any unhandled exception
    }

    /**
     * Helper method to assert start and end of parsing. Also works around a current race-condition where exception
     * could be dropped.
     */
    private static void assertStartEnd(final TestRead read) throws IOException, InterruptedException {
        // this call blocks on iterator: we need to wait for "parsing" to actually start
        assertEquals("START", read.next().get(0).getStringValue(), "Expected START of parsing test indicator");
        // give go for parser to throw exception
        read.throwNow();
        // this handles a race condition that currently exists in exception propagation between read and parser thread
        Thread.sleep(500);
        assertNull(read.next(), "Expected end to be reached");
    }

    private abstract static class TestRead extends ExcelRead {

        static final RandomAccessible<ExcelCell> START =
            RandomAccessibleUtils.createFromArray(new ExcelCell(KNIMECellType.STRING, "START"));

        /**
         * Latch used in test to coordinate when the parser should throw the exception.
         */
        private final CountDownLatch m_throwNow = new CountDownLatch(1);

        protected TestRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config,
            final Consumer<Map<String, Boolean>> sheetNamesConsumer) throws IOException {
            super(path, config, sheetNamesConsumer);
        }

        void throwNow() {
            // called by test method
            m_throwNow.countDown();
        }

        @Override
        public long getProgress() {
            throw new IllegalStateException("Should not have been called in test");
        }

        @Override
        public OptionalLong getMaxProgress() {
            throw new IllegalStateException("Should not have been called in test");
        }

        @Override
        public Set<Integer> getHiddenColumns() {
            throw new IllegalStateException("Should not have been called in test");
        }

        @Override
        protected ExcelParserRunnable createParser(final File localFile) throws IOException {
            return new ExcelParserRunnable(this, m_config) {

                @Override
                protected void parse() throws Throwable {
                    TestRead.this.addToQueue(START);
                    // Wait for latch to be initialized - handles race condition where parser thread
                    // starts before the outer anonymous TestRead instance is fully constructed
                    waitForCondition(() -> m_throwNow != null);
                    // fake blocking on I/O and throwing an exception (but wait not too long such that test stalls)
                    m_throwNow.await(30, TimeUnit.SECONDS);
                    throwException();
                }

                @Override
                protected void closeResources() throws IOException {
                    // no-op
                }
            };
        }

        protected abstract void throwException() throws Exception;
    }

    @SuppressWarnings("serial")
    static final class UnhandledException extends RuntimeException {
    }

    /**
     * Waits for a condition to become true, checking periodically with a 10-second timeout.
     *
     * @param condition the condition to wait for
     */
    private static void waitForCondition(final Callable<Boolean> condition) throws Exception {
        final long startTime = System.currentTimeMillis();
        final long timeoutMillis = 10_000;
        while (!condition.call()) {
            if (System.currentTimeMillis() - startTime > timeoutMillis) {
                throw new IllegalStateException("Timeout after 10 seconds waiting for condition");
            }
            Thread.sleep(10); // NOSONAR
        }
    }

}
