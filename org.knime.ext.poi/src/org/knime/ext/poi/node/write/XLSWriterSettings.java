/*
 * ------------------------------------------------------------------ *
 * This source code, its documentation and all appendant files
 * are protected by copyright law. All rights reserved.
 *
 * Copyright, 2003 - 2010
 * University of Konstanz, Germany
 * Chair for Bioinformatics and Information Mining (Prof. M. Berthold)
 * and KNIME GmbH, Konstanz, Germany
 *
 * You may not modify, publish, transmit, transfer or sell, reproduce,
 * create derivative works from, distribute, perform, display, or in
 * any way exploit any of the content, in whole or in part, except as
 * otherwise expressly permitted in writing by the copyright owner or
 * as specified in the license file distributed with this product.
 *
 * If you have any questions please contact the copyright holder:
 * website: www.knime.org
 * email: contact@knime.org
 * ---------------------------------------------------------------------
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
