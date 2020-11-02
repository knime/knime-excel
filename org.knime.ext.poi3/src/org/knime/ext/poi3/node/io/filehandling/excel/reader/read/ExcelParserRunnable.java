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
 *   Oct 22, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;

import org.knime.core.node.NodeLogger;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessible;
import org.knime.filehandling.core.node.table.reader.randomaccess.RandomAccessibleUtils;

/**
 * Abstract implementation of a {@link Runnable} that parses Excel files and adds the parsed rows as
 * {@link RandomAccessible} to the queue of the provided {@link ExcelRead}.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public abstract class ExcelParserRunnable implements Runnable {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelParserRunnable.class);

    private final ExcelRead m_read;

    /** True if 15 digits precision is used (that's what Excel is using). */
    protected final boolean m_use15DigitsPrecision;

    /**
     * Constructor.
     *
     * @param read the {@link ExcelRead}
     * @param config the config
     */
    public ExcelParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config) {
        m_read = read;
        m_use15DigitsPrecision = config.getReaderSpecificConfig().isUse15DigitsPrecision();
    }

    @Override
    public void run() {
        LOGGER.debug("Thread parsing an Excel spreadsheet started.");
        try {
            parse();
            m_read.addToQueue(ExcelRead.POISON_PILL);
            LOGGER.debug("Thread parsing an Excel spreadsheet finished successfully.");
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            closeAndSwallowExceptions();
        } catch (Throwable e) { // NOSONAR we want to catch any Throwable
            m_read.setThrowable(e);
            // cannot do anything with the IO exceptions besides logging
            closeAndSwallowExceptions();
        }
    }

    private void closeAndSwallowExceptions() {
        try {
            m_read.close();
        } catch (IOException ioe) {
            LOGGER.debug(ioe);
        }
    }

    /**
     * Start parsing.
     *
     * @throws Throwable if any error occurs
     */
    protected abstract void parse() throws Throwable; // NOSONAR throw anything and treat it in the main thread

    /**
     * Adds the {@link RandomAccessible} to the blocking queue of the runnable.
     *
     * @param randomAccessible the {@link RandomAccessible}
     */
    protected void addToQueue(final RandomAccessible<ExcelCell> randomAccessible) {
        try {
            m_read.addToQueue(randomAccessible);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            LOGGER.debug("Thread parsing an Excel spreadsheet interrupted.");
            // throw a runtime exception as the interrupting the thread does not stop it
            throw new ParsingInterruptedException();
        }
    }

    /**
     * Adds the given number of empty {@link RandomAccessible}s to the blocking queue of the runnable.
     *
     * @param numMissingsRows the number of missing rows to add
     */
    protected void outputEmptyRows(final int numMissingsRows) {
        for (int i = 0; i < numMissingsRows; i++) {
            addToQueue(RandomAccessibleUtils.createFromArrayUnsafe());
        }
    }

    /**
     * Exception to be thrown when the thread that parses the sheet should be interrupted.
     *
     * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
     */
    private static class ParsingInterruptedException extends RuntimeException {

        private static final long serialVersionUID = 1L;

    }

}
