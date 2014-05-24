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
 *   Mar 15, 2007 (ohl): created
 */
package org.knime.ext.poi.node.write;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;

/**
 * Holds the settings for the XLSWriter.
 *
 * @author ohl, University of Konstanz
 */
public class XLSWriterSettings {

    private static final String CFGKEY_XCELLOFFSET = "xCellOffset";

    private static final String CFGKEY_YCELLOFFSET = "yCellOffset";

    private static final String CFGKEY_WRITECOLHDR = "writeColHdr";

    private static final String CFGKEY_WRITEROWHDR = "writeRowHdr";

    private static final String CFGKEY_FILENAME = "filename";

    private static final String CFGKEY_SHEETNAME = "sheetname";

    private static final String CFGKEY_MISSINGPATTERN = "missingPattern";

    private static final String CFG_OVERWRITE_OK = "overwrite_ok";

    private int m_xCellOffset;

    private int m_yCellOffset;

    private boolean m_writeColHeader;

    private boolean m_writeRowID;

    private String m_filename;

    private String m_sheetname;

    private String m_missingPattern;

    private boolean m_overwriteOK;

    /**
     * Creates a new settings object with default settings but no filename.
     */
    public XLSWriterSettings() {
        m_xCellOffset = 0;
        m_yCellOffset = 0;
        m_writeColHeader = false;
        m_writeRowID = false;
        m_filename = null;
        m_sheetname = null;
        m_missingPattern = null;
        m_overwriteOK = false;
    }

    /**
     * Creates a new object with the setting values read from the specified
     * settings object.
     *
     * @param settings containing the values for the writer's settings.
     * @throws InvalidSettingsException if the specified settings object
     *             contains incorrect settings.
     */
    public XLSWriterSettings(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        m_xCellOffset = settings.getInt(CFGKEY_XCELLOFFSET);
        m_yCellOffset = settings.getInt(CFGKEY_YCELLOFFSET);
        m_writeColHeader = settings.getBoolean(CFGKEY_WRITECOLHDR);
        m_writeRowID = settings.getBoolean(CFGKEY_WRITEROWHDR);
        m_filename = settings.getString(CFGKEY_FILENAME);
        m_sheetname = settings.getString(CFGKEY_SHEETNAME);
        m_missingPattern = settings.getString(CFGKEY_MISSINGPATTERN);
        // option added for KNIME 2.0, we use "true" as default in cases
        // where this option is not present (old KNIME 1.x flows) in order
        // to mimic the old functionality
        m_overwriteOK = settings.getBoolean(CFG_OVERWRITE_OK, true);
    }

    /**
     * Saves the current values into the specified object.
     *
     * @param settings to write the values to.
     */
    public void saveSettingsTo(final NodeSettingsWO settings) {
        settings.addInt(CFGKEY_XCELLOFFSET, m_xCellOffset);
        settings.addInt(CFGKEY_YCELLOFFSET, m_yCellOffset);
        settings.addBoolean(CFGKEY_WRITECOLHDR, m_writeColHeader);
        settings.addBoolean(CFGKEY_WRITEROWHDR, m_writeRowID);
        settings.addString(CFGKEY_FILENAME, m_filename);
        settings.addString(CFGKEY_SHEETNAME, m_sheetname);
        settings.addString(CFGKEY_MISSINGPATTERN, m_missingPattern);
        settings.addBoolean(CFG_OVERWRITE_OK, m_overwriteOK);
    }

    /**
     * @return the filename
     */
    public String getFilename() {
        return m_filename;
    }

    /**
     * @param filename the filename to set
     */
    public void setFilename(final String filename) {
        m_filename = filename;
    }

    /**
     * @return the sheet name
     */
    public String getSheetname() {
        return m_sheetname;
    }

    /**
     * @param sheetname the sheet name to set, if null, the table name will be
     *            used as sheet name
     */
    public void setSheetname(final String sheetname) {
        m_sheetname = sheetname;
    }

    /**
     * @return the writeColHeader
     */
    public boolean writeColHeader() {
        return m_writeColHeader;
    }

    /**
     * @param writeColHeader the writeColHeader to set
     */
    public void setWriteColHeader(final boolean writeColHeader) {
        m_writeColHeader = writeColHeader;
    }

    /**
     * @return the writeRowID
     */
    public boolean writeRowID() {
        return m_writeRowID;
    }

    /**
     * @param writeRowID the writeRowID to set
     */
    public void setWriteRowID(final boolean writeRowID) {
        m_writeRowID = writeRowID;
    }

    /**
     * @return the xCellOffset
     */
    public int getXCellOffset() {
        return m_xCellOffset;
    }

    /**
     * @param cellOffset the xCellOffset to set
     */
    public void setXCellOffset(final int cellOffset) {
        m_xCellOffset = cellOffset;
    }

    /**
     * @return the yCellOffset
     */
    public int getYCellOffset() {
        return m_yCellOffset;
    }

    /**
     * @param cellOffset the yCellOffset to set
     */
    public void setYCellOffset(final int cellOffset) {
        m_yCellOffset = cellOffset;
    }

    /**
     * @param missingPattern the missingPattern to set. If null, no datasheet
     *            cell will be created.
     */
    public void setMissingPattern(final String missingPattern) {
        m_missingPattern = missingPattern;
    }

    /**
     * @return the missingPattern
     */
    public String getMissingPattern() {
        return m_missingPattern;
    }

    /** Set the overwrite ok property.
     * @param overwriteOK The property. */
    public void setOverwriteOK(final boolean overwriteOK) {
        m_overwriteOK = overwriteOK;
    }

    /** @return the overwrite ok property. */
    public boolean getOverwriteOK() {
        return (m_overwriteOK);
    }

}
