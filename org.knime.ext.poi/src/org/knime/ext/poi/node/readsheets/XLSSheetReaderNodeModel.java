/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by 
 *  University of Konstanz, Germany and
 *  KNIME GmbH, Konstanz, Germany
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
package org.knime.ext.poi.node.readsheets;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeCreationContext;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.ext.poi.POIActivator;
import org.knime.ext.poi.node.read2.XLSTableSettings;

/**
 * @author Patrick Winter, KNIME.com, Zurich, Switzerland
 */
public class XLSSheetReaderNodeModel extends NodeModel {

    private XLSSheetReaderSettings m_settings = new XLSSheetReaderSettings();

    /**
     * Creates the node model.
     */
    public XLSSheetReaderNodeModel() {
        super(0, 1);
        POIActivator.mkTmpDirRW_Bug3301();
    }

    /**
     * Creates the node model with the given context.
     *
     * @param context The context of this node instance
     */
    XLSSheetReaderNodeModel(final NodeCreationContext context) {
        this();
        m_settings.setFileLocation(context.getUrl().toString());
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
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
        m_settings = XLSSheetReaderSettings.load(settings);
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
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
        // empty
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        if (m_settings != null) {
            m_settings.save(settings);
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        XLSSheetReaderSettings.load(settings);
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
        return new DataTableSpec[]{createOutSpec()};
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData, final ExecutionContext exec)
            throws Exception {
        Workbook wb = getWorkbook(m_settings.getFileLocation());
        List<String> sheets = getSheetNames(wb);
        DataTableSpec spec = createOutSpec();
        BufferedDataContainer outContainer = exec.createDataContainer(spec);
        for (int i = 0; i < sheets.size(); i++) {
            outContainer.addRowToTable(new DefaultRow("Row" + i, new StringCell(sheets.get(i))));
        }
        outContainer.close();
        return new BufferedDataTable[]{outContainer.getTable()};
    }

    /**
     * Creates the specs of the output table.
     *
     * @return Specs of the output table.
     */
    private DataTableSpec createOutSpec() {
        return new DataTableSpec(new DataColumnSpecCreator("Sheet", StringCell.TYPE).createSpec());
    }

    /**
     * Returns the names of the sheets contained in the specified workbook.
     *
     * @param wb The workbook
     * @return a list with sheet names
     */
    private static ArrayList<String> getSheetNames(final Workbook wb) {
        ArrayList<String> names = new ArrayList<String>();
        for (int sIdx = 0; sIdx < wb.getNumberOfSheets(); sIdx++) {
            names.add(wb.getSheetName(sIdx));
        }
        return names;
    }

    /**
     * Loads a workbook from the file system.
     *
     * @param path Path to the workbook
     * @return The workbook
     * @throws IOException If the file could not be accessed
     * @throws InvalidFormatException If the file does not have the right format
     */
    private static Workbook getWorkbook(final String path) throws IOException, InvalidFormatException {
        Workbook workbook = null;
        InputStream in = null;
        try {
            in = XLSTableSettings.getBufferedInputStream(path);
            workbook = WorkbookFactory.create(in);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e2) {
                    // ignore
                }
            }
        }
        return workbook;
    }

}
