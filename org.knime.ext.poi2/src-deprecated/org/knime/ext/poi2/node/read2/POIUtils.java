/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME GmbH, Konstanz, Germany
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
 * ---------------------------------------------------------------------
 *
 * History
 *   11 Sep 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read2;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.util.FileUtil;
import org.knime.core.util.Pair;
import org.knime.ext.poi2.node.read2.KNIMEXSSFSheetXMLHandler.DataType;
import org.knime.ext.poi2.node.read2.KNIMEXSSFSheetXMLHandler.KNIMESheetContentsHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * Utility methods for accessing xls/xlsx files.
 *
 * @author Gabor Bakos
 */
@Deprecated
class POIUtils {
    /**
     * Checks whether the sheet is empty or not.
     */
    static class IsEmpty implements KNIMESheetContentsHandler {
        private boolean m_isEmpty = true;

        /**
         * {@inheritDoc}
         */
        @Override
        public void cell(final String cellReference, final String formattedValue, final XSSFComment comment) {
            m_isEmpty = false;
            throw new POIUtils.StopProcessing();
        }

        /**
         * @return the isEmpty
         */
        public boolean isEmpty() {
            return m_isEmpty;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void startRow(final int rowNum) {
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void endRow(final int rowNum) {
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void headerFooter(final String text, final boolean isHeader, final String tagName) {
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void nextCellType(final DataType type) {
        }

    }

    /**
     * Stops streaming visiting when this {@link RuntimeException} is thrown.
     */
    static final class StopProcessing extends RuntimeException {
        private static final long serialVersionUID = -6707877454679192967L;

    }

    /**
     * Combines column types. Combines only identical values, no further processing, keeps formulae.
     */
    static class ColumnTypeCombinator {
        private SortedMap<Integer, Pair<ActualDataType, Integer>> m_currentTypes = new TreeMap<>();

        ColumnTypeCombinator(final int firstRow, final Pair<ActualDataType, Integer> firstPair) {
            m_currentTypes.put(firstRow, firstPair);
        }

        private ActualDataType canonical(final ActualDataType type) {
            return type;
        }

        private Optional<ActualDataType> compatibility(final ActualDataType last, final ActualDataType next) {
            return last == next ? Optional.of(last) : Optional.empty();
        }


        void  combine(
            final int rowIndex0,
            final ActualDataType type) {
            CheckUtils.checkArgument(rowIndex0 >= 0, "Row index should be non-negative: " + rowIndex0);
            ActualDataType canonical = canonical(type);
            CheckUtils.checkState(m_currentTypes != null && !m_currentTypes.isEmpty(), "Current type is empty");

            Integer lastKey = m_currentTypes.lastKey();
            CheckUtils.checkState(lastKey < rowIndex0,
                "Cannot process a map with rows higher than the current " + lastKey + " >= " + rowIndex0);
            Pair<ActualDataType, Integer> lastValue = m_currentTypes.get(lastKey);
            CheckUtils.checkState(lastValue.getSecond() < rowIndex0,
                "Cannot process a map with rows right interval higher than the current " + lastValue.getSecond() +
                " >= " + rowIndex0);
            if (lastValue.getSecond() + 1 == rowIndex0) {//No missing rows are in-between
                Optional<ActualDataType> compatibility = compatibility(lastValue.getFirst(), canonical);
                if (compatibility.isPresent()) {
                    m_currentTypes.put(lastKey, Pair.create(compatibility.get(), rowIndex0));
                } else {
                    m_currentTypes.put(rowIndex0, Pair.create(canonical, rowIndex0));
                }
            } else {//Missings are between
                Optional<ActualDataType> compatibility = compatibility(lastValue.getFirst(), ActualDataType.MISSING);
                if (compatibility.isPresent()) {
                    Optional<ActualDataType> missingCompatibility = compatibility(ActualDataType.MISSING, canonical);
                    if (missingCompatibility.isPresent()) {
                        m_currentTypes.put(lastKey, Pair.create(compatibility.get(), rowIndex0));
                    } else {
                        m_currentTypes.put(lastKey, Pair.create(compatibility.get(), rowIndex0 - 1));
                        m_currentTypes.put(rowIndex0, Pair.create(canonical, rowIndex0));
                    }
                } else {
                    m_currentTypes.put(lastValue.getSecond() + 1, Pair.create(ActualDataType.MISSING, rowIndex0 - 1));
                    m_currentTypes.put(rowIndex0, Pair.create(canonical, rowIndex0));
                }
            }
        }

        /**
         * @return the currentTypes
         */
        public SortedMap<Integer, Pair<ActualDataType, Integer>> getCurrentTypes() {
            return new TreeMap<>(m_currentTypes);
        }
    }

    /**
     * Hidden constructor.
     */
    private POIUtils() {
    }

    /**
     * Returns the names of the sheets contained in the specified xls file.
     *
     * @param wb The workbook
     * @return an array with sheet names
     */
    public static ArrayList<String> getSheetNames(final Workbook wb) {

        ArrayList<String> names = new ArrayList<String>();
        for (int sIdx = 0; sIdx < wb.getNumberOfSheets(); sIdx++) {
            names.add(wb.getSheetName(sIdx));
        }
        return names;
    }

    /**
     * Finds out the name of the first sheet in the workbook.
     *
     * @param wBook the workbook to examine.
     * @return the name of the first sheet
     */
    public static String getFirstSheetNameWithData(final Workbook wBook) {
        String result = null;
        for (int sIdx = 0; sIdx < wBook.getNumberOfSheets(); sIdx++) {
            String sheetName = wBook.getSheetName(sIdx);
            if (sIdx == 0) {
                // return the first sheet, in case there is no data in the book
                result = sheetName;
            }
            Sheet sheet = wBook.getSheet(sheetName);
            int maxRowIdx = POIUtils.getLastRowIdx(sheet);
            int minRowIdx = sheet.getFirstRowNum();
            if (minRowIdx >= 0 && maxRowIdx >= 0 && minRowIdx <= maxRowIdx) {
                return sheetName;
            }
        }
        return result;

    }

    /**
     * @param A character between {@code A} and {@code Z}.c
     * @return Its numeric value (from {@code A = 1}).
     */
    static int toNumber(final char c) {
        assert c <= 'Z';
        assert c > '@';
        return c - '@';
    }

    /**
     * Returns the names of the sheets contained in the specified file.
     *
     * @param wb The workbook
     * @return an array with sheet names
     * @throws IOException Problem reading.
     * @throws InvalidFormatException
     */
    public static List<String> getSheetNames(final XSSFReader wb) throws InvalidFormatException, IOException {

        final List<String> names = new ArrayList<String>();
        for (final SheetIterator sheetsData = (SheetIterator)wb.getSheetsData(); sheetsData.hasNext();) {
            try (InputStream unused = sheetsData.next()) {
                names.add(sheetsData.getSheetName());
            }
        }
        return names;
    }

    /**
     * Finds out the name of the first sheet in the workbook.
     *
     * @param wBook the workbook to examine.
     * @param strings the strings of the workbook.
     * @return the name of the first sheet
     * @throws SAXException Problem reading.
     * @throws IOException Problem reading.
     * @throws OpenXML4JException Problem reading.
     * @throws ParserConfigurationException Problem reading.
     */
    public static String getFirstSheetNameWithData(final XSSFReader wBook, final ReadOnlySharedStringsTable strings)
        throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        String result = null;
        StylesTable styles = wBook.getStylesTable();
        SheetIterator iter = (SheetIterator)wBook.getSheetsData();
        while (iter.hasNext()) {
            try (final InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                if (result == null) {
                    result = sheetName;
                }
                final IsEmpty isEmpty = new IsEmpty();
                try {
                    processSheet(styles, strings, isEmpty, stream);
                } catch (final POIUtils.StopProcessing discard) {
                    //Ignore, this just means it is not empty.
                }
                if (!isEmpty.isEmpty()) {
                    return sheetName;
                }
            }
        }
        return result;

    }

    /**
     * Loads a workbook from the file system.
     *
     * @param path Path to the workbook
     * @return The workbook or null if it could not be loaded
     * @throws IOException Problem reading.
     * @throws InvalidFormatException Problem reading.
     * @throws RuntimeException the underlying POI library also throws other kind of exceptions
     */
    public static Workbook getWorkbook(final String path) throws IOException, InvalidFormatException {
        Workbook workbook = null;
        InputStream in = null;
        try {
            in = XLSUserSettings.getBufferedInputStream(path);
            // This should be the only place in the code where a workbook gets loaded
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

    /**
     * Works around the weirdness of the POI/XLS API.
     *
     * @param sheet the sheet to examine
     * @return the number of rows in the sheet minus one.
     */
    public static int getLastRowIdx(final Sheet sheet) {
        int rowMaxIdx = sheet.getLastRowNum();
        if (rowMaxIdx == 0) {
            if (sheet.getPhysicalNumberOfRows() == 0) {
                return -1;
            }
        }
        return rowMaxIdx;
    }

    /**
     * Computes the number of the column represented by the specified label.
     *
     * @param columnLabel string with the column's label. Must contain only capital letters.
     * @return the index of the column represented by the specified string ({@code 0}-based).
     * @throws NullPointerException if the specified label is null.
     * @throws IllegalArgumentException if the specified label is empty or contains anything else but capital letters.
     */
    static int getColumnIndex(final String columnLabel) {
        if (columnLabel.isEmpty()) {
            return -1;
        }
        if (!columnLabel.matches("[A-Z]+")) {
            throw new IllegalArgumentException("Not a valid column label: '" + columnLabel + "'");
        }

        int result = 0;
        for (int d = 0; d < columnLabel.length(); d++) {
            int dig = toNumber(columnLabel.charAt(columnLabel.length() - d - 1));
            result += dig * (int)Math.round(Math.pow(26, d));
        }

        return result - 1;
    }

    /**
     * Parses and shows the content of one sheet using the specified styles and shared-strings tables.
     *
     * @param styles The styles table.
     * @param strings The strings.
     * @param sheetInputStream The sheet's input stream.
     */
    private static void processSheet(final StylesTable styles, final ReadOnlySharedStringsTable strings,
        final SheetContentsHandler sheetHandler, final InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        DataFormatter formatter = new KNIMEDataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            org.xml.sax.ContentHandler handler =
                new XSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    /**
     * @param str A non-{@code null} {@link String}.
     * @return Checks whether it is an integer or not.
     */
    public static boolean isInteger(final String str) {
        try {
            //Integer.parseInt(str);
            double d = Double.parseDouble(str);
            return Double.isFinite(d) && d == (int)d;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * @param input A {@link String} value.
     * @return The parsed int in case it is a column reference. (Does not allow empty)
     * @throws InvalidSettingsException In case it is not a column reference.
     */
    static int oneBasedColumnNumberChecked(final String input) throws InvalidSettingsException {
        if (input == null || input.isEmpty()) {
            throw new InvalidSettingsException("please enter a column number.");
        }
        return oneBasedColumnNumber(input);
    }

    /**
     * @param input A possibly {@code null} or empty column reference.
     * @return The {@code 1}=based column reference ({@code 0} for empty).
     * @throws InvalidSettingsException Not a column reference.
     */
    static int oneBasedColumnNumber(final String input) throws InvalidSettingsException {
        if (input == null || input.isEmpty()) {
            return 0;
        }
        int i;
        try {
            i = Integer.parseInt(input);
        } catch (NumberFormatException nfe) {
            try {
                i = getColumnIndex(input.toUpperCase());
                // we want a number here (not an index)...
                i++;
            } catch (IllegalArgumentException iae) {
                throw new InvalidSettingsException(
                    "not a valid column number (enter a number or " + "'A', 'B', ..., 'Z', 'AA', etc.)");
            }
        }
        if (i <= 0) {
            throw new InvalidSettingsException("column number must be larger than zero");
        }
        return i;
    }

    /**
     * Converts a {@code 1}-based column reference to {@link String}.
     * @param input A positive int.
     * @return The column name (starting from {@code A}).
     * @throws InvalidSettingsException Wrong number.
     */
    static String oneBasedColumnNumberChecked(final int input) throws InvalidSettingsException {
        if (input <= 0) {
            throw new InvalidSettingsException("please enter a column number (greater than 0).");
        }
        return oneBasedColumnNumber(input);
    }

    /**
     *
     * @param input (between {@code 1} nad {@code 16384}). (For values below it returns {@code null}.)
     * @return {@code null} in case {@code input} is not valid.
     * @see #oneBasedColumnNumberChecked(int)
     */
    static String oneBasedColumnNumber(final int input) {
        if (input <= 0) {
            return null;
        }
        //assert input >= 0 : "Column index (1-based): " + input;
        if (input <= 26 /*Z*/) {
            return String.valueOf((char)('A' + (input - 1)));
        }
        int rem = (input - 1) % 26, div = (input - 1) / 26;
        return oneBasedColumnNumber(div) + String.valueOf((char)('A' + rem));
    }

    /**
     * Opens and returns a new buffered input stream on the passed location. The
     * location could either be a filename or a URL.
     *
     * @param location a filename or a URL
     * @param timeOutInSeconds ...
     * @return a new opened buffered input stream.
     * @throws IOException
     * @throws InvalidSettingsException
     */
    static InputStream openInputStream(
        final String location, final int timeOutInSeconds) throws IOException, InvalidSettingsException {
        return FileUtil.openInputStream(location, 1000 * timeOutInSeconds);
    }

    /**
     * Checks whether the type map contains only missing values in the column.
     *
     * @param columnMap The column's type information.
     * @param settings The user settings to specify what is considered missing value.
     * @return {@code true} iff all values are missing/empty.
     */
    static boolean isAllEmpty(final SortedMap<Integer, Pair<ActualDataType, Integer>> columnMap,
        final XLSUserSettings settings) {
        if (columnMap.isEmpty()) {
            return true;
        }
        int firstRow = settings.getFirstRow0(), lastRow = settings.getLastRow0();

        for (Entry<Integer, Pair<ActualDataType, Integer>> entry : columnMap.entrySet()) {
            if (firstRow >= 0 && entry.getValue().getSecond() < firstRow) {
                continue;
            }
            //Skip column header
            if (settings.getHasColHeaders() && entry.getKey() == settings.getColHdrRow0()
                && entry.getValue().getSecond() == settings.getColHdrRow0()) {
                continue;
            }
            if (lastRow >= 0 && entry.getKey() > lastRow) {
                break;
            }
            ActualDataType dataType = entry.getValue().getFirst();
            if (ActualDataType.isMissing(dataType)
                || (!settings.getUseErrorPattern() && ActualDataType.isError(dataType))) {
                continue;
            }
            return false;
        }
        return true;
    }

    /**
     * Provides names for {@code null} {@code colHdrs}.
     * @param settings The user settings.
     * @param skippedCols The skipped columns.
     * @param colHdrs The known column headers.
     */
    static void fillEmptyColHeaders(final XLSUserSettings settings, final Set<Integer> skippedCols,
        final String[] colHdrs) {
        //TODO should we use something like ValueUniquifier here?
        HashSet<String> names = new HashSet<String>();
        for (int i = 0; i < colHdrs.length; i++) {
            if (colHdrs[i] != null) {
                names.add(colHdrs[i]);
            }
        }

        int xlOffset = 0;

        for (int i = 0; i < colHdrs.length; i++) {
            while (skippedCols.contains(xlOffset)) {
                xlOffset++;
            }
            if (colHdrs[i] == null) {
                String colName = colNameForIndex(i);
                if (settings.getKeepXLColNames()) {
                    colName = oneBasedColumnNumber(Math.max(1, settings.getFirstColumn()) + xlOffset);
                }
                colHdrs[i] = getUniqueName(colName, names);
                names.add(colHdrs[i]);
            }
            xlOffset++;
        }
    }

    /**
     * @param i An index.
     * @return Default column name.
     */
    static String colNameForIndex(final int i) {
        return "Col" + i;
    }

    /**
     * @param name Proposed name.
     * @param names Existing names (this method does not modify it, you should).
     * @return Unique name for a column.
     */
    static String getUniqueName(final String name, final Set<String> names) {
        int cnt = 2;
        String unique = name;
        while (names.contains(unique)) {
            unique = name + "_" + cnt++;
        }
        return unique;
    }
}
