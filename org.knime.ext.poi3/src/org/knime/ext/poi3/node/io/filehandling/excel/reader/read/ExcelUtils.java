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
 *   Nov 3, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * Utility class for Excel nodes.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelUtils {

    private ExcelUtils() {
        // Hide constructor, utils class
    }

    /**
     * Returns the name of the first sheet with data, i.e., non-empty sheet. If there is no sheet with data, the first
     * one is returned.
     *
     * @param sheetNames the map of sheet names and booleans that indicate if a sheet is the first with data
     *
     * @return the name of the first sheet with data or first sheet if all empty
     */
    public static String getFirstSheetWithDataOrFirstIfAllEmpty(final Map<String, Boolean> sheetNames) {
        if (sheetNames.isEmpty()) {
            return ""; // should never happen; still, handle gracefully to prevent UI crashes
        }
        final String firstSheet = sheetNames.keySet().iterator().next();
        for (final Entry<String, Boolean> sheetNameEntry : sheetNames.entrySet()) {
            if (sheetNameEntry.getValue().booleanValue()) {
                return sheetNameEntry.getKey();
            }
        }
        return firstSheet;
    }

    /**
     * Returns a map that contains the names of the sheets contained in the specified {@link Workbook} as keys and
     * whether it is the first non-empty sheet as value.
     *
     * @param workbook the workbook
     * @return the map of sheet names and whether a sheet is the first with data
     */
    public static Map<String, Boolean> getSheetNames(final Workbook workbook) {
        final Map<String, Boolean> sheetNames = new LinkedHashMap<>(); // LinkedHashMap to retain order
        boolean nonEmptySheetFound = false;
        for (final Sheet sheet : workbook) {
            if (nonEmptySheetFound) {
                sheetNames.put(sheet.getSheetName(), false);
            } else {
                final boolean isEmpty = isEmpty(sheet);
                nonEmptySheetFound = !isEmpty;
                sheetNames.put(sheet.getSheetName(), !isEmpty);
            }
        }
        return sheetNames;
    }

    /**
     * Returns a map that contains the names of the sheets contained in the file read by the specified
     * {@link XSSFReader} as keys and whether it is the first non-empty sheet as value.
     *
     * @param xmlReader the xml reader
     * @param reader the xssf reader
     * @param sharedStrings the shared strings table
     * @return the map of sheet names and whether a sheet is the first with data
     * @throws InvalidFormatException
     * @throws IOException
     * @throws SAXException
     */
    public static Map<String, Boolean> getSheetNames(final XMLReader xmlReader, final XSSFReader reader,
        final SharedStrings sharedStrings) throws InvalidFormatException, IOException, SAXException {
        final Map<String, Boolean> sheetNames = new LinkedHashMap<>(); // LinkedHashMap to retain order
        boolean nonEmptySheetFound = false;
        xmlReader.setContentHandler(
            new XSSFSheetXMLHandler(reader.getStylesTable(), sharedStrings, new IsEmpty(), new DataFormatter(), false));
        final SheetIterator sheetsData = (SheetIterator)reader.getSheetsData();
        while (sheetsData.hasNext()) {
            try (final InputStream inputStream = sheetsData.next()) {
                if (nonEmptySheetFound) {
                    sheetNames.put(sheetsData.getSheetName(), false);
                } else {
                    final boolean sheetEmpty = isSheetEmpty(xmlReader, inputStream);
                    sheetNames.put(sheetsData.getSheetName(), !sheetEmpty);
                    nonEmptySheetFound = !sheetEmpty;
                }
            }
        }
        return sheetNames;
    }

    private static boolean isSheetEmpty(final XMLReader xmlReader, final InputStream inputStream)
        throws IOException, SAXException {
        try {
            xmlReader.parse(new InputSource(inputStream));
            return true;
        } catch (ParsingInterruptedException e) { // NOSONAR, exception is expected and handled
            // not empty
            return false;
        }
    }

    /**
     * We need to iterate over all rows and columns as a cell could be blank but have a style set. In such cases,
     * {@link Sheet#getLastRowNum()} would not return 0 and a row would not be {@code null}.
     */
    private static boolean isEmpty(final Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            final Row row = sheet.getRow(i);
            if (row != null && !isRowEmpty(row)) {
                return false;
            }
        }
        return true;
    }

    private static boolean isRowEmpty(final Row row) {
        for (int j = 0; j < row.getLastCellNum(); j++) {
            final Cell cell = row.getCell(j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null) {
                return false;
            }
        }
        return true;
    }

    /**
     * Checks whether the sheet is empty or not.
     */
    static class IsEmpty implements SheetContentsHandler {

        @Override
        public void cell(final String cellReference, final String formattedValue, final XSSFComment comment) {
            throw new ParsingInterruptedException();
        }

        @Override
        public void startRow(final int rowNum) {
            // do nothing
        }

        @Override
        public void endRow(final int rowNum) {
            // do nothing
        }

        @Override
        public void headerFooter(final String text, final boolean isHeader, final String tagName) {
            // do nothing
        }

    }

    /**
     * Exception to be thrown when the thread that parses the sheet should be interrupted.
     *
     * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
     */
    static class ParsingInterruptedException extends RuntimeException {

        private static final long serialVersionUID = 1L;

    }

}
