/*
 * ------------------------------------------------------------------------
 *  Copyright by KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
 * -------------------------------------------------------------------
 *
 * History
 *   Apr 8, 2009 (ohl): created
 */
package org.knime.ext.poi2.node.read2;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Locale;
import java.util.concurrent.ExecutionException;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeCreationContext;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.util.FileUtil;
import org.knime.ext.poi2.POIActivator;
import org.xml.sax.SAXException;

/**
 * Xls(x) reader node's {@link NodeModel}.
 *
 * @author Peter Ohl, KNIME.com, Zurich, Switzerland
 * @author Gabor Bakos
 */
public class XLSReaderNodeModel extends NodeModel {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(XLSReaderNodeModel.class);

    private XLSUserSettings m_settings = new XLSUserSettings();

    private DataTableSpec m_dts = null;

    private String m_dtsSettingsID = null;

    /**
     *
     */
    public XLSReaderNodeModel() {
        super(0, 1);
        POIActivator.mkTmpDirRW_Bug3301();
    }

    XLSReaderNodeModel(final NodeCreationContext context) {
        this();
        m_settings.setFileLocation(context.getUrl().toString());
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        /*
         * This is a special "deal" for the reader: The reader, if previously
         * executed, has data at it's output - even if the file that was read
         * doesn't exist anymore. In order to warn the user that the data cannot
         * be recreated we check here if the file exists and set a warning
         * message if it doesn't.
         */
        String fName = m_settings.getFileLocation();
        if (fName == null || fName.isEmpty()) {
            return;
        }

        try {
            new URL(fName);
            // don't check URLs - don't open a stream.
            return;
        } catch (MalformedURLException mue) {
            // continue on a file
        }
        File location = new File(fName);

        if (!location.canRead() || location.isDirectory()) {
            setWarningMessage("The file '" + location.getAbsolutePath() + "' can't be accessed anymore!");
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_settings = XLSUserSettings.load(settings);
        try {
            String settingsID = settings.getString(XLSReaderNodeDialog.XLS_CFG_ID_FOR_TABLESPEC);
            if (!m_settings.getID().equals(settingsID)) {
                throw new InvalidSettingsException("IDs don't match");
            }
            NodeSettingsRO dtsConfig = settings.getNodeSettings(XLSReaderNodeDialog.XLS_CFG_TABLESPEC);
            m_dts = DataTableSpec.load(dtsConfig);
            m_dtsSettingsID = settingsID;
        } catch (InvalidSettingsException ise) {
            LOGGER.debug("No DTS saved in settings");
            // it's optional - if it's not saved we create it later
            m_dts = null;
            m_dtsSettingsID = null;
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void reset() {
        // empty
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec)
        throws IOException, CanceledExecutionException {
        // empty
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        if (m_settings != null) {
            m_settings.save(settings);
            if (m_dts != null) {
                settings.addString(XLSReaderNodeDialog.XLS_CFG_ID_FOR_TABLESPEC, m_dtsSettingsID);
                m_dts.save(settings.addConfig(XLSReaderNodeDialog.XLS_CFG_TABLESPEC));
            }
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        XLSUserSettings.load(settings);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final DataTableSpec[] inSpecs) throws InvalidSettingsException {

        if (m_settings == null) {
            throw new InvalidSettingsException("Node not configured.");
        }

        String errMsg = m_settings.getStatus(true);
        if (errMsg != null) {
            throw new InvalidSettingsException(errMsg);
        }

        // make sure the DTS still fits the settings
        if (!m_settings.getID().equals(m_dtsSettingsID)) {
            m_dts = null;
        }
        if (m_dts == null) {
            // this could take a while.
            LOGGER.debug("Building DTS during configure...");
            XLSUserSettings settings = XLSUserSettings.clone(m_settings);
            String sheetName = settings.getSheetName();
            String loc = settings.getFileLocation();
            try {
                sheetName = settings(settings, sheetName, loc);
            } catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException e1) {
                throw new InvalidSettingsException(e1);
            }
            settings.setSheetName(sheetName);
            CachedExcelTable table;
            try (InputStream is = FileUtil.openInputStream(loc);
                    InputStream stream = new BufferedInputStream(is);
                    CountingInputStream countingStream = new CountingInputStream(stream)) {
                table = isXlsx() && !m_settings.isReevaluateFormulae()
                    ? CachedExcelTable.fillCacheFromXlsxStreaming(m_settings.getFileLocation(), countingStream,
                        sheetName, Locale.ENGLISH, new ExecutionMonitor(), null).get()
                    : CachedExcelTable.fillCacheFromDOM(m_settings.getFileLocation(), countingStream, sheetName,
                        Locale.ENGLISH, m_settings.isReevaluateFormulae(), new ExecutionMonitor(), null).get();
                m_dts = table.createDataTable(settings, null).getDataTableSpec();
            } catch (InterruptedException | ExecutionException | RuntimeException | IOException e) {
                throw new InvalidSettingsException(e);
            }
            m_dtsSettingsID = m_settings.getID();
        }
        return new DataTableSpec[]{m_dts};
    }

    /**
     * @param e
     * @return
     */
    String message(final Exception e) {
        String execMsg = e.getMessage();
        if (execMsg == null) {
            execMsg = e.getClass().getSimpleName();
        }
        return execMsg;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData, final ExecutionContext exec)
        throws Exception {
        // Execute is an isolated call so do not keep the workbook in memory
        XLSUserSettings settings = XLSUserSettings.clone(m_settings);
        String sheetName = settings.getSheetName();
        String loc = settings.getFileLocation();
        sheetName = settings(settings, sheetName, loc);
        try (InputStream is = FileUtil.openInputStream(loc, settings.getTimeoutInMilliseconds());
                CountingInputStream countingStream = new CountingInputStream(is)) {
            CachedExcelTable table = isXlsx() && !settings.isReevaluateFormulae()
                ? CachedExcelTable.fillCacheFromXlsxStreaming(loc, countingStream, sheetName, Locale.ENGLISH,
                    exec.createSubExecutionContext(.9), null).get()
                : CachedExcelTable.fillCacheFromDOM(loc, countingStream, sheetName, Locale.ENGLISH,
                    settings.isReevaluateFormulae(), exec.createSubExecutionContext(.9), null).get();
            return new BufferedDataTable[]{exec.createBufferedDataTable(table.createDataTable(settings, null), exec)};
        }
    }

    /**
     * @param settings
     * @param sheetName
     * @param loc
     * @return
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws InvalidFormatException
     */
    private String settings(final XLSUserSettings settings, String sheetName, final String loc)
        throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, InvalidFormatException {
        if (sheetName == null || XLSReaderNodeDialog.FIRST_SHEET.equals(sheetName)) {
            if (isXlsx()) {
                try (final BufferedInputStream stream = POIUtils.getBufferedInputStream(loc);
                        final OPCPackage pack = OPCPackage.open(stream)) {
                    sheetName =
                        POIUtils.getFirstSheetNameWithData(new XSSFReader(pack), new ReadOnlySharedStringsTable(pack));
                }
            } else {
                sheetName = POIUtils.getFirstSheetNameWithData(POIUtils.getWorkbook(loc));
            }
            settings.setSheetName(sheetName);
        }
        return sheetName;
    }

    final boolean isXlsx() {
        return isXlsx(m_settings.getFileLocation());
    }

    /**
     * @param fileName The file name/url
     * @return Whether its name ends with {@code .xsls} (case insensitive).
     */
    static final boolean isXlsx(final String fileName) {
        return fileName.toLowerCase().endsWith(".xlsx");
    }
}
