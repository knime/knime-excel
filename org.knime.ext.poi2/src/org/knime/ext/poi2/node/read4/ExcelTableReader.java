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
 *   Sep 25, 2019 (julian): created
 */
package org.knime.ext.poi2.node.read4;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutionException;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.base.node.io.filehandling.FilesToDataTableReader;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTable;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.streamable.RowOutput;
import org.xml.sax.SAXException;

/**
 * Specific implementation of {@link FilesToDataTableReader} to process {@link Path Path(s)} leading to Excel files.
 *
 * @author Julian Bunzel, KNIME GmbH, Berlin, Germany
 */
class ExcelTableReader implements FilesToDataTableReader {

    private final XLSUserSettings m_settings;

    private String m_sheetName;

    private boolean m_sheetNameSet = false;

    long m_curRow = 0L - 1;

    long m_totalRows = 0L;

    long m_lastNoOfRows = 0;

    private DataTable m_firstTable;

    private DataTableSpec m_spec = null;

    private boolean m_processedTableFromSpecCreation = false;

    private final ValueUniquifier m_uniquifier = new ValueUniquifier();

    /**
     * Creates a new instance of {@code ExcelTableReader}.
     *
     * @param settings {@code XLSUserSettings} containing all necessary information how to process the files
     * @throws InvalidSettingsException
     */
    ExcelTableReader(final XLSUserSettings settings) throws InvalidSettingsException {
        m_settings = XLSUserSettings.clone(settings);
        m_sheetName = m_settings.getSheetName();
    }

    private DataTable createDataTable(final Path path, final ExecutionMonitor parseExec)
        throws IOException, InterruptedException, ExecutionException, InvalidSettingsException, SAXException,
        OpenXML4JException, ParserConfigurationException {
        // Validate and set sheet name based on first path handled by this instance of ExcelTableReader
        if (!m_sheetNameSet) {
            m_sheetName = settings(path, m_settings, m_sheetName);
            m_settings.setSheetName(m_sheetName);
            m_sheetNameSet = true;
        }
        try (final InputStream is = Files.newInputStream(path)) {
            final CachedExcelTable table = isXlsx(path) && !m_settings.isReevaluateFormulae()
                ? CachedExcelTable.fillCacheFromXlsxStreaming(path, is, m_sheetName, Locale.ENGLISH, parseExec, null)
                    .get()
                : CachedExcelTable.fillCacheFromDOM(path, is, m_sheetName, Locale.ENGLISH,
                    m_settings.isReevaluateFormulae(), parseExec, null).get();

            final DataTable dt = table.createDataTable(path, m_settings, null, m_curRow, m_totalRows, m_uniquifier);
            if (!m_processedTableFromSpecCreation && (m_firstTable == null)) {
                m_firstTable = dt;
                m_spec = dt.getDataTableSpec();
            } else {
                validateTableSpecs(path, dt);
            }
            m_lastNoOfRows = table.lastRow();
            m_totalRows += m_lastNoOfRows;
            return dt;
        }
    }

    private void validateTableSpecs(final Path path, final DataTable table) {
        if (m_spec != null && !table.getDataTableSpec().equalStructure(m_spec)) {
            throw new RuntimeException("Table created from file '" + path.getFileName().toString()
                + "' has different structure than previously read file(s)");
        }
    }


    @Override
    public void pushRowsToOutput(final Path path, final RowOutput output, final ExecutionContext exec)
        throws Exception {
        if ((m_firstTable != null) && !m_processedTableFromSpecCreation) {
            exec.setMessage("Writing cashed rows to table");
            pushRows(output, m_firstTable, exec);
            exec.setMessage("Writing finished");
            m_firstTable = null;
        } else {
            final DataTable table = createDataTable(path, exec.createSubProgress(0.5));
            exec.setMessage("Writing cashed rows to table");
            pushRows(output, table, exec);
            exec.setMessage("Writing finished");
            if (!m_processedTableFromSpecCreation) {
                m_processedTableFromSpecCreation = true;
                m_firstTable = null;
            }
        }
    }

    private void pushRows(final RowOutput output, final DataTable table, final ExecutionMonitor exec)
            throws InterruptedException, CanceledExecutionException {
        double counter = 0;
        for (final DataRow r : table) {
            final long rowCount = (long)counter;
            exec.setProgress(counter++/m_lastNoOfRows, () -> "Writing row " + rowCount);
            exec.checkCanceled();
            output.push(r);
            m_curRow += 1L;
        }
    }

    @Override
    public DataTableSpec createDataTableSpec(final List<Path> paths, final ExecutionMonitor exec)
            throws InvalidSettingsException {
        if (m_firstTable != null) {
            return m_firstTable.getDataTableSpec();
        }
        try {
            return createDataTable(paths.get(0), exec).getDataTableSpec();
        } catch (final Exception e) {
            throw new InvalidSettingsException(e);
        }
    }

    private static String settings(final Path path, final XLSUserSettings settings, String sheetName)
        throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {

        if ((sheetName == null) || XLSReaderNodeDialog.FIRST_SHEET.equals(sheetName)) {
            if (isXlsx(path)) {
                try (InputStream stream = Files.newInputStream(path); final OPCPackage pack = OPCPackage.open(stream)) {
                    sheetName = POIUtils.getFirstSheetNameWithData(new XSSFReader(pack),
                        new ReadOnlySharedStringsTable(pack, false));
                }
            } else {
                sheetName = POIUtils.getFirstSheetNameWithData(POIUtils.getWorkbook(path));
            }
            settings.setSheetName(sheetName);
        }
        return sheetName;
    }

    /**
     * Checks whether the path leads to an .xlsx file.
     *
     * @return True, if .xlsx
     */
    static final boolean isXlsx(final Path path) {
        return isXlsx(path.getFileName().toString());
    }

    /**
     * @param fileName The file name/url
     * @return Whether its name ends with {@code .xsls} (case insensitive).
     */
    static final boolean isXlsx(final String fileName) {
        return fileName.toLowerCase().endsWith(".xlsx");
    }
}
