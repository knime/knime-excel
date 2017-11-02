/*
 * ------------------------------------------------------------------------
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
package org.knime.ext.poi2.node.write2;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;

/**
 * Holds the settings for the XLSWriter.
 *
 * @author ohl, University of Konstanz
 */
public class XLSWriter2Settings {

    private static final String CFGKEY_XCELLOFFSET = "xCellOffset";

    private static final String CFGKEY_YCELLOFFSET = "yCellOffset";

    private static final String CFGKEY_WRITECOLHDR = "writeColHdr";

    private static final String CFGKEY_WRITEROWHDR = "writeRowHdr";

    private static final String CFGKEY_FILENAME = "filename";

    private static final String CFGKEY_SHEETNAME = "sheetname";

    private static final String CFGKEY_MISSINGPATTERN = "missingPattern";

    private static final String CFG_OVERWRITE_OK = "overwrite_ok";

    private static final String CFG_OPEN_FILE = "open_file";

    private static final String CFG_FILE_MUST_EXIST = "file_must_exist";

    private static final String CFG_DO_NOT_OVERWRITE_SHEET = "do_not_overwrite_sheet";

    private static final String CFG_LANDSCAPE = "landscape";

    private static final String CFG_AUTOSIZE = "autosize";

    private static final String CFG_PAPER_SIZE = "paper_size";

    private static final String CFG_WRITE_MISSING_VALUES = "write_missing_values";

    private static final String CFG_EVALUATE_FORMULA = "evaluate_formula";

    private int m_xCellOffset;

    private int m_yCellOffset;

    private boolean m_writeColHeader;

    private boolean m_writeRowID;

    private String m_filename;

    private String m_sheetname;

    private String m_missingPattern;

    private boolean m_overwriteOK;

    private boolean m_openFile;

    private boolean m_fileMustExist;

    private boolean m_doNotOverwriteSheet;

    private boolean m_landscape;

    private boolean m_autosize;

    private short m_paperSize;

    private boolean m_writeMissingValues;

    private boolean m_evaluateFormula;

    /**
     * Creates a new settings object with default settings but no filename.
     */
    public XLSWriter2Settings() {
        m_xCellOffset = 0;
        m_yCellOffset = 0;
        m_writeColHeader = false;
        m_writeRowID = false;
        m_filename = null;
        m_sheetname = null;
        m_missingPattern = null;
        m_overwriteOK = false;
        m_openFile = false;
        m_fileMustExist = false;
        m_doNotOverwriteSheet = false;
        m_landscape = false;
        m_paperSize = PrintSetup.LETTER_PAPERSIZE;
        m_writeMissingValues = false;
        m_evaluateFormula = false;
    }

    /**
     * Creates a new settings object with default settings but no filename.
     *
     * @param spec Specification containing the input tables name
     */
    public XLSWriter2Settings(final DataTableSpec spec) {
        this();
        m_sheetname = spec.getName();
    }

    /**
     * Creates a new object with the setting values read from the specified settings object.
     *
     * @param settings containing the values for the writer's settings.
     * @throws InvalidSettingsException if the specified settings object contains incorrect settings.
     */
    public XLSWriter2Settings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_xCellOffset = settings.getInt(CFGKEY_XCELLOFFSET);
        m_yCellOffset = settings.getInt(CFGKEY_YCELLOFFSET);
        m_writeColHeader = settings.getBoolean(CFGKEY_WRITECOLHDR);
        m_writeRowID = settings.getBoolean(CFGKEY_WRITEROWHDR);
        m_filename = settings.getString(CFGKEY_FILENAME);
        m_sheetname = settings.getString(CFGKEY_SHEETNAME);
        if (m_sheetname == null) {
            throw new InvalidSettingsException("Please specify a sheet name.");
        }
        m_missingPattern = settings.getString(CFGKEY_MISSINGPATTERN);
        // option added for KNIME 2.0, we use "true" as default in cases
        // where this option is not present (old KNIME 1.x flows) in order
        // to mimic the old functionality
        m_overwriteOK = settings.getBoolean(CFG_OVERWRITE_OK, true);
        m_openFile = settings.getBoolean(CFG_OPEN_FILE, false);
        m_fileMustExist = settings.getBoolean(CFG_FILE_MUST_EXIST, false);
        m_doNotOverwriteSheet = settings.getBoolean(CFG_DO_NOT_OVERWRITE_SHEET, false);
        m_landscape = settings.getBoolean(CFG_LANDSCAPE, false);
        m_autosize = settings.getBoolean(CFG_AUTOSIZE, false);
        m_paperSize = settings.getShort(CFG_PAPER_SIZE, PrintSetup.LETTER_PAPERSIZE);
        m_writeMissingValues = settings.getBoolean(CFG_WRITE_MISSING_VALUES, true);
        m_evaluateFormula = settings.getBoolean(CFG_EVALUATE_FORMULA, false);
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
        settings.addBoolean(CFG_OPEN_FILE, m_openFile);
        settings.addBoolean(CFG_FILE_MUST_EXIST, m_fileMustExist);
        settings.addBoolean(CFG_DO_NOT_OVERWRITE_SHEET, m_doNotOverwriteSheet);
        settings.addBoolean(CFG_LANDSCAPE, m_landscape);
        settings.addBoolean(CFG_AUTOSIZE, m_autosize);
        settings.addShort(CFG_PAPER_SIZE, m_paperSize);
        settings.addBoolean(CFG_WRITE_MISSING_VALUES, m_writeMissingValues);
        settings.addBoolean(CFG_EVALUATE_FORMULA, m_evaluateFormula);
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
     * @param sheetname the sheet name to set, if null, the table name will be used as sheet name
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
     * @param missingPattern the missingPattern to set. If null, no datasheet cell will be created.
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

    /**
     * Set the overwrite ok property.
     *
     * @param overwriteOK The property.
     */
    public void setOverwriteOK(final boolean overwriteOK) {
        m_overwriteOK = overwriteOK;
    }

    /** @return the overwrite ok property. */
    public boolean getOverwriteOK() {
        return (m_overwriteOK);
    }

    /**
     * @param openFile Should the file be opened after execution
     */
    public void setOpenFile(final boolean openFile) {
        m_openFile = openFile;
    }

    /**
     * @return openFile
     */
    public boolean getOpenFile() {
        return m_openFile;
    }

    /**
     * @param fileMustExist the fileMustExist to set
     */
    public void setFileMustExist(final boolean fileMustExist) {
        m_fileMustExist = fileMustExist;
    }

    /**
     * @return the fileMustExist
     */
    public boolean getFileMustExist() {
        return m_fileMustExist;
    }

    /**
     * @param doNotOverwriteSheet the doNotOverwriteSheet to set
     */
    public void setDoNotOverwriteSheet(final boolean doNotOverwriteSheet) {
        m_doNotOverwriteSheet = doNotOverwriteSheet;
    }

    /**
     * @return the doNotOverwriteSheet
     */
    public boolean getDoNotOverwriteSheet() {
        return m_doNotOverwriteSheet;
    }

    /**
     * @param landscape the landscape to set
     */
    public void setLandscape(final boolean landscape) {
        m_landscape = landscape;
    }

    /**
     * @return the landscape
     */
    public boolean getLandscape() {
        return m_landscape;
    }

    /**
     * @param autosize the autosize to set
     */
    public void setAutosize(final boolean autosize) {
        m_autosize = autosize;
    }

    /**
     * @return the autosize
     */
    public boolean getAutosize() {
        return m_autosize;
    }

    /**
     * @param paperSize the paperSize to set
     */
    public void setPaperSize(final short paperSize) {
        m_paperSize = paperSize;
    }

    /**
     * @return the paperSize
     */
    public short getPaperSize() {
        return m_paperSize;
    }

    /**
     * @param writeMissingValues the writeMissingValues to set
     * @since 2.11
     */
    public void setWriteMissingValues(final boolean writeMissingValues) {
        m_writeMissingValues = writeMissingValues;
    }

    /**
     * @return the writeMissingValues
     * @since 2.11
     */
    public boolean getWriteMissingValues() {
        return m_writeMissingValues;
    }

    /**
     * @return the writeMissingValues
     * @since 3.1.1
     */
    public boolean evaluateFormula() {
        return m_evaluateFormula;
    }

    /**
     * @param evaluateFormula
     * @since 3.1.1
     */
    public void setEvaluateFormula(final boolean evaluateFormula) {
        m_evaluateFormula = evaluateFormula;
    }


}
