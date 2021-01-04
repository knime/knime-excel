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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.node.workflow.CredentialsProvider;
import org.knime.filehandling.core.node.table.reader.config.ReaderSpecificConfig;

/**
 * An implementation of {@link ReaderSpecificConfig} class for Excel table reader.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelTableReaderConfig implements ReaderSpecificConfig<ExcelTableReaderConfig> {

    static final String CFG_PASSWORD = "password";

    /** The default String to be inserted when an error cell is encountered. */
    static final String DEFAULT_FORMULA_ERROR_PATTERN = "#XL_EVAL_ERROR#";

    private boolean m_useRawSettings = false;

    private boolean m_use15DigitsPrecision = true;

    private SheetSelection m_sheetSelection = SheetSelection.FIRST;

    private String m_sheetName = "";

    private int m_sheetIdx = 0;

    private boolean m_skipHiddenCols = true;

    private boolean m_skipHiddenRows = true;

    private boolean m_reevaluateFormulas = false;

    private boolean m_replaceEmptyStringsWithMissings = true;

    private FormulaErrorHandling m_formulaErrorHandling = FormulaErrorHandling.PATTERN;

    private String m_errorPattern = DEFAULT_FORMULA_ERROR_PATTERN;

    private AreaOfSheetToRead m_areaOfSheetToRead = AreaOfSheetToRead.ENTIRE;

    private String m_readFromCol = "A";

    private String m_readToCol = "";

    private String m_readFromRow = "1";

    private String m_readToRow = "";

    private RowIDGeneration m_rowIdGeneration = RowIDGeneration.GENERATE;

    private String m_rowIDCol = "A";

    private boolean m_zipBombDetection = true;

    private SettingsModelAuthentication m_authenticationSettingsModel =
        new SettingsModelAuthentication(CFG_PASSWORD, AuthenticationType.NONE);

    private CredentialsProvider m_credentialsProvider;

    /**
     * Constructor.
     */
    ExcelTableReaderConfig() {
    }

    private ExcelTableReaderConfig(final ExcelTableReaderConfig toCopy) {
        setUse15DigitsPrecision(toCopy.isUse15DigitsPrecision());
        setSheetSelection(toCopy.getSheetSelection());
        setSheetName(toCopy.getSheetName());
        setSheetIdx(toCopy.getSheetIdx());
        setSkipHiddenCols(toCopy.isSkipHiddenCols());
        setSkipHiddenRows(toCopy.isSkipHiddenRows());
        setReevaluateFormulas(toCopy.isReevaluateFormulas());
        setFormulaErrorHandling(toCopy.getFormulaErrorHandling());
        setErrorPattern(toCopy.getErrorPattern());
        setAreaOfSheetToRead(toCopy.getAreaOfSheetToRead());
        setReadFromCol(toCopy.getReadFromCol());
        setReadToCol(toCopy.getReadToCol());
        setReadFromRow(toCopy.getReadFromRow());
        setReadToRow(toCopy.getReadToRow());
        setRowIdGeneration(toCopy.getRowIdGeneration());
        setRowIDCol(toCopy.getRowIDCol());
        setUseRawSettings(toCopy.isUseRawSettings());
        setReplaceEmptyStringsWithMissings(toCopy.isReplaceEmptyStringsWithMissings());
        setAuthenticationSettingsModel(toCopy.getAuthenticationSettingsModel());
        setCredentialsProvider(toCopy.getCredentialsProvider());
        setZipBombDetection(toCopy.isZipBombDetection());
    }

    @Override
    public ExcelTableReaderConfig copy() {
        return new ExcelTableReaderConfig(this);
    }

    /**
     * @return the use15DigitsPrecision
     */
    public boolean isUse15DigitsPrecision() {
        return m_use15DigitsPrecision;
    }

    /**
     * @param use15DigitsPrecision the use15DigitsPrecision to set
     */
    void setUse15DigitsPrecision(final boolean use15DigitsPrecision) {
        m_use15DigitsPrecision = use15DigitsPrecision;
    }

    /**
     * @return the sheetSelection
     */
    public SheetSelection getSheetSelection() {
        return m_sheetSelection;
    }

    /**
     * @param sheetSelection the sheetSelection to set
     */
    void setSheetSelection(final SheetSelection sheetSelection) {
        m_sheetSelection = sheetSelection;
    }

    /**
     * @return the selectedSheetName
     */
    public String getSheetName() {
        return m_sheetName;
    }

    /**
     * @param selectedSheetName the selectedSheetName to set
     */
    void setSheetName(final String selectedSheetName) {
        m_sheetName = selectedSheetName;
    }

    /**
     * @return the selectedSheetIdx
     */
    public int getSheetIdx() {
        return m_sheetIdx;
    }

    /**
     * @param selectedSheetIdx the selectedSheetIdx to set
     */
    void setSheetIdx(final int selectedSheetIdx) {
        m_sheetIdx = selectedSheetIdx;
    }

    /**
     * @return the skipHiddenCols
     */
    public boolean isSkipHiddenCols() {
        return m_skipHiddenCols;
    }

    /**
     * @param skipHiddenCols the skipHiddenCols to set
     */
    void setSkipHiddenCols(final boolean skipHiddenCols) {
        m_skipHiddenCols = skipHiddenCols;
    }

    /**
     * @return the skipHiddenRows
     */
    public boolean isSkipHiddenRows() {
        return m_skipHiddenRows;
    }

    /**
     * @param skipHiddenRows the skipHiddenRows to set
     */
    void setSkipHiddenRows(final boolean skipHiddenRows) {
        m_skipHiddenRows = skipHiddenRows;
    }

    /**
     * @return the reevaluateFormulas
     */
    public boolean isReevaluateFormulas() {
        return m_reevaluateFormulas;
    }

    /**
     * @param reevaluateFormulas the reevaluateFormulas to set
     */
    void setReevaluateFormulas(final boolean reevaluateFormulas) {
        m_reevaluateFormulas = reevaluateFormulas;
    }

    /**
     * @return the formulaErrorHandling
     */
    public FormulaErrorHandling getFormulaErrorHandling() {
        return m_formulaErrorHandling;
    }

    /**
     * @param formulaErrorHandling the formulaErrorHandling to set
     */
    void setFormulaErrorHandling(final FormulaErrorHandling formulaErrorHandling) {
        m_formulaErrorHandling = formulaErrorHandling;
    }

    /**
     * @return the errorPattern
     */
    public String getErrorPattern() {
        return m_errorPattern;
    }

    /**
     * @param errorPattern the errorPattern to set
     */
    void setErrorPattern(final String errorPattern) {
        m_errorPattern = errorPattern;
    }

    /**
     * @return the readFromCol
     */
    public String getReadFromCol() {
        return m_readFromCol;
    }

    /**
     * @param readFromCol the readFromCol to set
     */
    void setReadFromCol(final String readFromCol) {
        m_readFromCol = readFromCol;
    }

    /**
     * @return the readToCol
     */
    public String getReadToCol() {
        return m_readToCol;
    }

    /**
     * @param readToCol the readToCol to set
     */
    void setReadToCol(final String readToCol) {
        m_readToCol = readToCol;
    }

    /**
     * @return the readFromRow
     */
    public String getReadFromRow() {
        return m_readFromRow;
    }

    /**
     * @param readFromRow the readFromRow to set
     */
    void setReadFromRow(final String readFromRow) {
        m_readFromRow = readFromRow;
    }

    /**
     * @return the readToRow
     */
    public String getReadToRow() {
        return m_readToRow;
    }

    /**
     * @param readToRow the readToRow to set
     */
    void setReadToRow(final String readToRow) {
        m_readToRow = readToRow;
    }

    /**
     * @return the areaOfSheetToRead
     */
    public AreaOfSheetToRead getAreaOfSheetToRead() {
        return m_areaOfSheetToRead;
    }

    /**
     * @param areaOfSheetToRead the areaOfSheetToRead to set
     */
    void setAreaOfSheetToRead(final AreaOfSheetToRead areaOfSheetToRead) {
        m_areaOfSheetToRead = areaOfSheetToRead;
    }

    /**
     * @return the rowIDCol
     */
    public String getRowIDCol() {
        return m_rowIDCol;
    }

    /**
     * @param rowIDCol the rowIDCol to set
     */
    void setRowIDCol(final String rowIDCol) {
        m_rowIDCol = rowIDCol;
    }

    /**
     * @return the rowIdGeneration
     */
    public RowIDGeneration getRowIdGeneration() {
        return m_rowIdGeneration;
    }

    /**
     * @param rowIdGeneration the rowIdGeneration to set
     */
    void setRowIdGeneration(final RowIDGeneration rowIdGeneration) {
        m_rowIdGeneration = rowIdGeneration;
    }

    /**
     * @return the useRawSettings
     */
    public boolean isUseRawSettings() {
        return m_useRawSettings;
    }

    /**
     * @param useRawSettings the useRawSettings to set
     */
    void setUseRawSettings(final boolean useRawSettings) {
        m_useRawSettings = useRawSettings;
    }

    /**
     * @return the replaceEmptyStringsWithMissings
     */
    public boolean isReplaceEmptyStringsWithMissings() {
        return m_replaceEmptyStringsWithMissings;
    }

    /**
     * @param replaceEmptyStringsWithMissings the replaceEmptyStringsWithMissings to set
     */
    void setReplaceEmptyStringsWithMissings(final boolean replaceEmptyStringsWithMissings) {
        m_replaceEmptyStringsWithMissings = replaceEmptyStringsWithMissings;
    }

    /**
     * @return the authenticationSettingsModel
     */
    public SettingsModelAuthentication getAuthenticationSettingsModel() {
        return m_authenticationSettingsModel;
    }

    /**
     * @param authenticationSettingsModel the authenticationSettingsModel to set
     */
    void setAuthenticationSettingsModel(final SettingsModelAuthentication authenticationSettingsModel) {
        m_authenticationSettingsModel = authenticationSettingsModel;
    }

    /**
     * @return the credentialsProvider
     */
    public CredentialsProvider getCredentialsProvider() {
        return m_credentialsProvider;
    }

    /**
     * @param credentialsProvider the credentialsProvider to set
     */
    void setCredentialsProvider(final CredentialsProvider credentialsProvider) {
        m_credentialsProvider = credentialsProvider;
    }

    /**
     * @return the zipBombDetection
     */
    public boolean isZipBombDetection() {
        return m_zipBombDetection;
    }

    /**
     * @param zipBombDetection the zipBombDetection to set
     */
    void setZipBombDetection(final boolean zipBombDetection) {
        m_zipBombDetection = zipBombDetection;
    }

}
