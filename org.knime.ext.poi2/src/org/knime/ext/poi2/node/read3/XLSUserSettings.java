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
 * -------------------------------------------------------------------
 *
 * History
 *   Apr 3, 2009 (ohl): created
 */
package org.knime.ext.poi2.node.read3;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.util.CheckUtils;

/**
 * Configuration settings for POI reading of xls(x) files.
 *
 * @author Peter Ohl, KNIME AG, Zurich, Switzerland
 * @author Gabor Bakos
 */
class XLSUserSettings {

    /** Key under which the file url is stored. */
    static final String CFG_XLS_LOCATION = "XLS_LOCATION";

    /**
     * Key for reevaluating formulae.
     */
    private static final String REEVALUATE_FORMULAE = "REEVALUATE_FORMULAE";

    private static final String TIMEOUT_IN_SECONDS = "TIMEOUT_IN_SECONDS";

    /** Default timeout in seconds. */
    static final int DEFAULT_TIMEOUT_IN_SECONDS = 1;

    /** Default value for the no preview setting. */
    static final boolean DEFAULT_NO_PREVIEW = false;

    private String m_fileLocation;

    private boolean m_readAllData;

    private int m_firstRow0;

    private int m_lastRow0;

    private int m_firstColumn0;

    private int m_lastColumn0;

    private String m_sheetName;

    private boolean m_hasColHeaders;

    private int m_colHdrRow0;

    private boolean m_hasRowHeaders;

    private int m_rowHdrCol0;

    private boolean m_indexContinuous;

    private boolean m_indexSkipJumps;

    private String m_missValuePattern;

    private boolean m_skipEmptyColumns;

    private boolean m_skipEmptyRows;

    private boolean m_skipHiddenColumns;

    private boolean m_keepXLColNames;

    private boolean m_uniquifyRowIDs;

    private String m_errorPattern;

    private boolean m_useErrorPattern;

    /** When {@code true} it uses DOM reading instead of streaming and reevaluates formulae. */
    private boolean m_reevaluateFormulae;

    private boolean m_noPreview = DEFAULT_NO_PREVIEW;

    private int m_timeoutInSeconds = DEFAULT_TIMEOUT_IN_SECONDS;

    static final boolean DEFAULT_REEVALUATE_FORMULAE = false;

    /** Default pattern for formula evaluation error StringCells */
    static final String DEFAULT_ERR_PATTERN = "#XL_EVAL_ERROR#";

    private static final String NO_PREVIEW = "NO_PREVIEW";

    /**
     * Constructs a new settings object. All values are uninitialized.
     */
    XLSUserSettings() {
        m_fileLocation = null;

        m_sheetName = null;

        m_readAllData = true;
        m_firstRow0 = -1;
        m_lastRow0 = -1;
        m_firstColumn0 = -1;
        m_lastColumn0 = -1;

        m_hasColHeaders = false;
        m_colHdrRow0 = 0;
        m_indexContinuous = true;
        m_indexSkipJumps = false;
        m_hasRowHeaders = false;
        m_rowHdrCol0 = 0;

        m_missValuePattern = null;

        m_skipEmptyColumns = true;
        m_skipEmptyRows = true;
        m_skipHiddenColumns = true;

        m_keepXLColNames = false;
        m_uniquifyRowIDs = false;

        m_useErrorPattern = true;
        m_errorPattern = DEFAULT_ERR_PATTERN;

        m_reevaluateFormulae = DEFAULT_REEVALUATE_FORMULAE;

        m_noPreview = DEFAULT_NO_PREVIEW;
    }

    /**
     * Saves the current values.
     *
     * @param settings object to write values into
     */
    final void save(final NodeSettingsWO settings) {
        CheckUtils.checkState(
            (m_hasRowHeaders && !m_indexContinuous && !m_indexSkipJumps)
                || (!m_hasRowHeaders && m_indexContinuous && !m_indexSkipJumps)
                || (!m_hasRowHeaders && !m_indexContinuous && m_indexSkipJumps),
            "Exactly one of generate row ids or table contains row ids in column should be selected!");
        settings.addString(CFG_XLS_LOCATION, m_fileLocation);

        settings.addString("SHEET_NAME", m_sheetName);

        settings.addBoolean("READ_ALL_DATA", m_readAllData);
        settings.addInt("FIRST_ROW", getFirstRow());
        settings.addInt("LAST_ROW", getLastRow());
        settings.addString("FIRST_COL", POIUtils.oneBasedColumnNumber(getFirstColumn()));
        settings.addString("LAST_COL", POIUtils.oneBasedColumnNumber(getLastColumn()));

        settings.addBoolean("HAS_COL_HDRS", m_hasColHeaders);
        settings.addInt("COL_HDR_ROW", getColHdrRow());
        settings.addBoolean("INDEX_CONTINUOUS", m_indexContinuous);
        settings.addBoolean("INDEX_JUMP_SKIPS", m_indexSkipJumps);
        settings.addBoolean("HAS_ROW_HDRS", m_hasRowHeaders);
        settings.addString("ROW_HDR_COL", POIUtils.oneBasedColumnNumber(getRowHdrCol()));

        settings.addString("MISS_VAL_PATTERN", m_missValuePattern);
        settings.addBoolean("SKIP_EMPTY_COLS", m_skipEmptyColumns);
        settings.addBoolean("SKIP_EMPTY_ROWS", m_skipEmptyRows);
        settings.addBoolean("SKIP_HIDDEN_COLS", m_skipHiddenColumns);

        settings.addBoolean("KEEP_XL_HDR", m_keepXLColNames);
        settings.addBoolean("UNIQUIFY_ROWID", m_uniquifyRowIDs);

        settings.addBoolean("FORMULA_ERR_USE_PATTERN", m_useErrorPattern);
        settings.addString("FORMULA_ERR_PATTERN", m_errorPattern);

        settings.addBoolean(REEVALUATE_FORMULAE, m_reevaluateFormulae);
        settings.addInt(TIMEOUT_IN_SECONDS, m_timeoutInSeconds);
        settings.addBoolean(NO_PREVIEW, m_noPreview);
    }

    /**
     * @return Simplified id of the settings.
     */
    final String getID() {
        StringBuilder id = new StringBuilder("SettingsID:");
        id.append(getID(m_fileLocation));
        id.append(getID(m_sheetName));
        id.append(getID(m_readAllData));
        id.append(getID(getFirstRow()));
        id.append(getID(getLastRow()));
        id.append(getID(getFirstColumn()));
        id.append(getID(getLastColumn()));
        id.append(getID(m_hasColHeaders));
        id.append(getID(getColHdrRow()));
        //id of m_indexSkipJumps, m_indexContinuous are not important, those are just for generated row keys.
        id.append(getID(m_hasRowHeaders));
        id.append(getID(getRowHdrCol()));
        id.append(getID(m_missValuePattern));
        id.append(getID(m_skipEmptyColumns));
        id.append(getID(m_skipEmptyRows));
        id.append(getID(m_skipHiddenColumns));
        id.append(getID(m_keepXLColNames));
        id.append(getID(m_uniquifyRowIDs));
        id.append(getID(m_useErrorPattern));
        id.append(getID(m_errorPattern));
        id.append(getID(m_reevaluateFormulae));
        id.append(getID(m_timeoutInSeconds));
        id.append(getID(m_noPreview));
        return id.toString();
    }

    /**
     * @param value A value.
     * @return Id of a {@link String} value.
     */
    private static String getID(final String value) {
        if (value == null) {
            return "-";
        }
        if (value.isEmpty()) {
            return ".";
        }
        return value;
    }
    /**
     * @param value A boolean value.
     * @return Id of the {@code value}.
     */
    private static String getID(final boolean value) {
        return value ? "1" : "0";
    }
    /**
     * @param value An int value.
     * @return Id of the {@code value}.
     */
    private static String getID(final int value) {
        return "" + value;
    }

    /**
     * Creates a new settings object with values from the settings passed.
     *
     * @param settings the values to store in the new object
     * @return a new settings object
     * @throws InvalidSettingsException if the stored settings are incorrect.
     */
    static final XLSUserSettings load(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        XLSUserSettings result = new XLSUserSettings();

        result.m_fileLocation = settings.getString(CFG_XLS_LOCATION);

        result.m_sheetName = settings.getString("SHEET_NAME");

        result.m_readAllData = settings.getBoolean("READ_ALL_DATA");
        result.setFirstColumn(POIUtils.oneBasedColumnNumber(settings.getString("FIRST_COL")));
        result.setLastColumn(POIUtils.oneBasedColumnNumber(settings.getString("LAST_COL")));
        result.setFirstRow(settings.getInt("FIRST_ROW"));
        result.setLastRow(settings.getInt("LAST_ROW"));

        result.m_hasColHeaders = settings.getBoolean("HAS_COL_HDRS");
        result.setColHdrRow(settings.getInt("COL_HDR_ROW"));
        result.m_hasRowHeaders = settings.getBoolean("HAS_ROW_HDRS");
        result.setRowHdrCol(POIUtils.oneBasedColumnNumber(settings.getString("ROW_HDR_COL")));
        try {
            result.m_indexContinuous = settings.getBoolean("INDEX_CONTINUOUS");
            result.m_indexSkipJumps = settings.getBoolean("INDEX_JUMP_SKIPS");
        } catch (InvalidSettingsException e) {
            //Added in 3.3.2
            if (result.m_hasRowHeaders) {
                result.m_indexContinuous = false;
                result.m_indexSkipJumps = false;
            } else {
                result.m_indexContinuous = false;
                result.m_indexSkipJumps = true;
            }
        }

        result.m_missValuePattern = settings.getString("MISS_VAL_PATTERN");
        result.m_skipEmptyColumns = settings.getBoolean("SKIP_EMPTY_COLS");
        result.m_skipEmptyRows = settings.getBoolean("SKIP_EMPTY_ROWS");
        result.m_skipHiddenColumns = settings.getBoolean("SKIP_HIDDEN_COLS");

        result.m_keepXLColNames = settings.getBoolean("KEEP_XL_HDR");
        result.m_uniquifyRowIDs = settings.getBoolean("UNIQUIFY_ROWID");

        // added in V2.2.3
        result.m_useErrorPattern =
                settings.getBoolean("FORMULA_ERR_USE_PATTERN", true);
        result.m_errorPattern =
                settings.getString("FORMULA_ERR_PATTERN", DEFAULT_ERR_PATTERN);

        result.m_reevaluateFormulae = settings.getBoolean(REEVALUATE_FORMULAE, DEFAULT_REEVALUATE_FORMULAE);
        result.m_timeoutInSeconds = settings.getInt(TIMEOUT_IN_SECONDS, DEFAULT_TIMEOUT_IN_SECONDS);
        result.m_noPreview = settings.getBoolean(NO_PREVIEW, DEFAULT_NO_PREVIEW);
        return result;
    }

    /**
     * Creates a deep copy of the passed settings object.
     *
     * @param orig the settings object to clone
     * @return a copy of the argument
     * @throws InvalidSettingsException normally not
     */
    static final XLSUserSettings clone(final XLSUserSettings orig)
            throws InvalidSettingsException {
        NodeSettings clone = new NodeSettings("clone");
        orig.save(clone);
        return XLSUserSettings.load(clone);
    }

    /**
     * Checks the settings and returns an error message - or null if everything
     * is alright.
     *
     * @return an error message or null if settings are okay
     */
    final String getStatus() {
        if (m_fileLocation == null || m_fileLocation.isEmpty()) {
            return "No file location specified";
        }
        if (m_timeoutInSeconds < 0) {
            return "Timeout should be non-negative! (0 means infinite)";
        }

        if (!m_readAllData) {
            if (m_firstColumn0 < 0) {
                return "'First column' index must be greater than one";
            }
            if (m_firstRow0 < 0) {
                return "'First row' index must be greater than one";
            }
            if (m_lastColumn0 >= 0 && m_lastColumn0 < m_firstColumn0) {
                return "Last column to read can't be "
                        + "smaller than the first column";
            }
            if (m_lastRow0 >= 0 && m_lastRow0 < m_firstRow0) {
                return "Last Row to read can't be smaller than the first row";
            }
        }
        if (m_hasColHeaders && m_colHdrRow0 < 0) {
            return "Row containing column headers is not specified";
        }
        if (m_hasRowHeaders && m_rowHdrCol0 < 0) {
            return "Column containing row IDs is not specified";
        }
        if (!(
            (m_hasRowHeaders && !m_indexContinuous && !m_indexSkipJumps)
                || (!m_hasRowHeaders && m_indexContinuous && !m_indexSkipJumps)
                || (!m_hasRowHeaders && !m_indexContinuous && m_indexSkipJumps))) {
            return "Exactly one of generate row ids or table contains row ids in column should be selected!";
        }

        return null;
    }

    /*
     * ---------------- setter and getter ------------------------------------
     */

    /**
     * @param fileLocation the fileLocation to set
     */
    final void setFileLocation(final String fileLocation) {
        m_fileLocation = fileLocation;
    }

    /**
     * @return the fileLocation
     */
    final String getFileLocation() {
        return m_fileLocation;
    }

    /**
     * @return the simple name of the file
     */
    final String getSimpleFilename() {
        String path = m_fileLocation;
        int idx = path.lastIndexOf('/');
        if (idx < 0 || idx >= path.length() - 2) {
            return path;
        } else {
            return path.substring(idx + 1);
        }
    }

    /**
     * @param sheetName name of sheet to read
     */
    final void setSheetName(final String sheetName) {
        m_sheetName = sheetName;
    }

    /**
     * @return the index of the sheet to read
     */
    final String getSheetName() {
        return m_sheetName;
    }

    /**
     * @return the firstRow ({@code 0}-based)
     */
    final int getFirstRow0() {
        return m_firstRow0;
    }

    /**
     * @return the firstRow ({@code 1}-based)
     */
    final int getFirstRow() {
        return m_firstRow0 + 1;
    }

    /**
     * @param firstRow the firstRow ({@code 0}-based) to set
     */
    final void setFirstRow0(final int firstRow) {
        m_firstRow0 = firstRow;
    }

    /**
     * @param firstRow the firstRow ({@code 1}-based) to set
     */
    final void setFirstRow(final int firstRow) {
        m_firstRow0 = firstRow - 1;
    }

    /**
     * @return the lastRow ({@code 0}-based)
     */
    final int getLastRow0() {
        return m_lastRow0;
    }

    /**
     * @return the lastRow ({@code 1}-based)
     */
    final int getLastRow() {
        return m_lastRow0 + 1;
    }

    /**
     * @param lastRow the lastRow ({@code 0}-based) to set
     */
    final void setLastRow0(final int lastRow) {
        m_lastRow0 = lastRow;
    }

    /**
     * @param lastRow the lastRow ({@code 1}-based) to set
     */
    final void setLastRow(final int lastRow) {
        m_lastRow0 = lastRow - 1;
    }

    /**
     * @return the firstColumn ({@code 0}-based)
     */
    final int getFirstColumn0() {
        return m_firstColumn0;
    }

    /**
     * @return the firstColumn ({@code 1}-based)
     */
    final int getFirstColumn() {
        return m_firstColumn0 + 1;
    }

    /**
     * @param firstColumn the firstColumn ({@code 1}-based) to set
     */
    final void setFirstColumn(final int firstColumn) {
        m_firstColumn0 = firstColumn - 1;
    }

    /**
     * @param firstColumn the firstColumn to set ({@code 0}-based)
     */
    final void setFirstColumn0(final int firstColumn) {
        m_firstColumn0 = firstColumn;
    }

    /**
     * @return the lastColumn ({@code 0}-based) (can be {@code -1})
     */
    final int getLastColumn0() {
        return m_lastColumn0;
    }

    /**
     * @return the lastColumn ({@code 1}-based) (can be {@code 0})
     */
    final int getLastColumn() {
        return m_lastColumn0 + 1;
    }

    /**
     * @param lastColumn the lastColumn ({@code 0}-based) to set (can be {@code -1})
     */
    final void setLastColumn0(final int lastColumn) {
        m_lastColumn0 = lastColumn;
    }

    /**
     * @param lastColumn the lastColumn ({@code 1}-based) to set (can be {@code 0})
     */
    final void setLastColumn(final int lastColumn) {
        m_lastColumn0 = lastColumn  - 1;
    }

    /**
     * @param skipHiddenColumns the skipHiddenColumns to set
     */
    final void setSkipHiddenColumns(final boolean skipHiddenColumns) {
        m_skipHiddenColumns = skipHiddenColumns;
    }

    /**
     * @return the skipHiddenColumns
     */
    final boolean getSkipHiddenColumns() {
        return m_skipHiddenColumns;
    }

    /**
     * @return true if empty columns are skipped
     */
    final boolean getSkipEmptyColumns() {
        return m_skipEmptyColumns;
    }

    /**
     * @param skipEmptyColumns the skipEmptyColumns to set
     */
    final void setSkipEmptyColumns(final boolean skipEmptyColumns) {
        m_skipEmptyColumns = skipEmptyColumns;
    }

    /**
     * @param skipEmptyRows the skipEmptyRows to set
     */
    final void setSkipEmptyRows(final boolean skipEmptyRows) {
        m_skipEmptyRows = skipEmptyRows;
    }

    /**
     * @return the skipEmptyRows
     */
    final boolean getSkipEmptyRows() {
        return m_skipEmptyRows;
    }

    /**
     * @return the colHdrRow ({@code 0}-based) (row of the column names)
     */
    final int getColHdrRow0() {
        return m_colHdrRow0;
    }

    /**
     * @return the colHdrRow ({@code 1}-based) (row of the column names)
     */
    final int getColHdrRow() {
        return m_colHdrRow0 + 1;
    }

    /**
     * @param colHdrRow the colHdrRow ({@code 0}-based) (row of the column names) to set
     */
    final void setColHdrRow0(final int colHdrRow) {
        m_colHdrRow0 = colHdrRow;
    }

    /**
     * @param colHdrRow the colHdrRow ({@code 1}-based) (row of the column names) to set
     */
    final void setColHdrRow(final int colHdrRow) {
        m_colHdrRow0 = colHdrRow - 1;
    }

    /**
     * @return the hasColHeaders (column names are specified)
     */
    final boolean getHasColHeaders() {
        return m_hasColHeaders;
    }

    /**
     * @param hasColHeaders the hasColHeaders (column names are specified) to set
     */
    final void setHasColHeaders(final boolean hasColHeaders) {
        m_hasColHeaders = hasColHeaders;
    }

    /**
     * @param hasRowHeaders the hasRowHeaders (row ids are specified) to set
     */
    final void setHasRowHeaders(final boolean hasRowHeaders) {
        m_hasRowHeaders = hasRowHeaders;
    }

    /**
     * @return the hasRowHeaders (row ids are specified)
     */
    final boolean getHasRowHeaders() {
        return m_hasRowHeaders;
    }

    /**
     * @param rowHdrCol the rowHdrCol ({@code 0}-based) (column of row ids) to set
     */
    final void setRowHdrCol0(final int rowHdrCol) {
        m_rowHdrCol0 = rowHdrCol;
    }

    /**
     * @param rowHdrCol the rowHdrCol ({@code 1}-based) (column of row ids) to set
     */
    final void setRowHdrCol(final int rowHdrCol) {
        m_rowHdrCol0 = rowHdrCol - 1;
    }

    /**
     * @return the rowHdrCol ({@code 0}-based) (column of row ids)
     */
    final int getRowHdrCol0() {
        return m_rowHdrCol0;
    }

    /**
     * @return the rowHdrCol ({@code 1}-based) (column of row ids)
     */
    final int getRowHdrCol() {
        return m_rowHdrCol0 + 1;
    }


    /**
     * @return the indexContinuous (the row index is continuous starting from {@code Row0} when {@code true})
     */
    final boolean isIndexContinuous() {
        return m_indexContinuous;
    }

    /**
     * @param indexContinuous the indexContinuous (the row index is continuous starting from {@code Row0} when
     *            {@code true}) to set
     */
    final void setIndexContinuous(final boolean indexContinuous) {
        m_indexContinuous = indexContinuous;
    }

    /**
     * @return the indexSkipJumps (row keys have jumps for missing rows)
     */
    final boolean isIndexSkipJumps() {
        return m_indexSkipJumps;
    }

    /**
     * @param indexSkipJumps the indexSkipJumps (row keys have jumps for missing rows) to set
     */
    final void setIndexSkipJumps(final boolean indexSkipJumps) {
        m_indexSkipJumps = indexSkipJumps;
    }

    /**
     * @return the keepXLColNames
     */
    final boolean getKeepXLColNames() {
        return m_keepXLColNames;
    }

    /**
     * @param keepXLColNames the keepXLColNames to set
     */
    final void setKeepXLNames(final boolean keepXLColNames) {
        m_keepXLColNames = keepXLColNames;
    }

    /**
     * @return the missValuePattern
     */
    final String getMissValuePattern() {
        return m_missValuePattern;
    }

    /**
     * @param missValuePattern the missValuePattern to set
     */
    final void setMissValuePattern(final String missValuePattern) {
        m_missValuePattern = missValuePattern;
    }

    /**
     * @return the readAllData
     */
    final boolean getReadAllData() {
        return m_readAllData;
    }

    /**
     * @param readAllData the readAllData to set
     */
    final void setReadAllData(final boolean readAllData) {
        m_readAllData = readAllData;
    }

    /**
     * @return the uniquifyRowIDs
     */
    final boolean getUniquifyRowIDs() {
        return m_uniquifyRowIDs;
    }

    /**
     * @param uniquifyRowIDs the uniquifyRowIDs to set
     */
    final void setUniquifyRowIDs(final boolean uniquifyRowIDs) {
        m_uniquifyRowIDs = uniquifyRowIDs;
    }

    /**
     * @see #setUseErrorPattern(boolean)
     * @param errorPattern the errorPattern that is inserted if a formula
     *            evaluation fails (only if the corresponding boolean flag is
     *            true).
     */
    final void setErrorPattern(final String errorPattern) {
        if (errorPattern == null) {
            m_errorPattern = "";
        } else {
            m_errorPattern = errorPattern;
        }
    }

    /**
     * @see #getUseErrorPattern()
     * @return the errorPattern. A StringCell with the pattern is inserted for
     *         the formula value in case the formula eval fails - and the
     *         corresponding flag is set true.
     */
    final String getErrorPattern() {
        return m_errorPattern;
    }

    /**
     * @see #setErrorPattern(String)
     * @param useErrorPattern if true the error pattern set is inserted (in a
     *            StringCell) if formula evaluations fail (causing the entire
     *            column to become a string column).
     */
    final void setUseErrorPattern(final boolean useErrorPattern) {
        m_useErrorPattern = useErrorPattern;
    }

    /**
     * @see XLSUserSettings#getErrorPattern()
     * @return true, if a StringCell is inserted when formula evaluation fails.
     *         Or false, if a missing cell is inserted instead.
     */
    final boolean getUseErrorPattern() {
        return m_useErrorPattern;
    }

    /**
     * @return the reevaluate formulae (DOM instead of stream reading) value
     */
    final boolean isReevaluateFormulae() {
        return m_reevaluateFormulae;
    }

    /**
     * @param reevaluateFormulae the reevaluate formulae (DOM instead of stream reading) to set
     */
    final void setReevaluateFormulae(final boolean reevaluateFormulae) {
        m_reevaluateFormulae = reevaluateFormulae;
    }

    /**
     * @return the timeoutInSeconds
     */
    final int getTimeoutInSeconds() {
        return m_timeoutInSeconds;
    }

    /**
     * @param timeoutInSeconds the timeoutInSeconds to set
     */
    final void setTimeoutInSeconds(final int timeoutInSeconds) {
        m_timeoutInSeconds = timeoutInSeconds;
    }

    /**
     * @return the noPreview
     */
    final boolean isNoPreview() {
        return m_noPreview;
    }

    /**
     * @param noPreview the noPreview to set
     */
    final void setNoPreview(final boolean noPreview) {
        m_noPreview = noPreview;
    }

    /**
     * Normalizes the settings.
     *
     * @param rawSettings Probably hard to use settings (but the sheet name exists).
     * @return Easier to use settings (with first/last columns set to "invalid"/special values for read all data.)
     * @throws InvalidSettingsException
     */
    static final XLSUserSettings normalizeSettings(final XLSUserSettings rawSettings) throws InvalidSettingsException {
        final XLSUserSettings settings = clone(rawSettings);
        if (settings.getReadAllData()) {
            settings.setLastRow0(-1);
            settings.setLastColumn0(-1);
            settings.setFirstRow0(-1);
            settings.setFirstColumn0(-1);
        }
        return settings;
    }
}
