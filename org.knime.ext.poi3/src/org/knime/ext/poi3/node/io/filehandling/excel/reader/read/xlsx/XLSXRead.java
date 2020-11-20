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
 *   Oct 20, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xlsx;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCellUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelParserRunnable;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xlsx.KNIMEXSSFSheetXMLHandler.AbstractKNIMESheetContentsHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xlsx.KNIMEXSSFSheetXMLHandler.KNIMEXSSFDataType;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * This class implements a read that uses the streaming API of Apache POI (eventmodel API) and can read xlsx and xlsm
 * file formats.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class XLSXRead extends ExcelRead {

    /**
     * Constructor.
     *
     * @param path the path of the file to read
     * @param config the Excel table read config
     * @throws IOException if an I/O exception occurs
     */
    public XLSXRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        super(path, config);
        // don't do any initializations here, super constructor will call #createParser(InputStream)
    }

    private OPCPackage m_opc;

    private CountingInputStream m_countingSheetStream;

    private long m_sheetSize;

    private Map<String, Boolean> m_sheetNames;

    private Set<Integer> m_hiddenColumns;

    @Override
    public ExcelParserRunnable createParser(final InputStream inputStream) throws IOException {
        try {
            m_opc = OPCPackage.open(inputStream);
            final XSSFReader xssfReader = new XSSFReader(m_opc);
            final XMLReader xmlReader = XMLHelper.newXMLReader();
            final ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(m_opc, false);

            m_sheetNames = ExcelUtils.getSheetNames(xmlReader, xssfReader, sharedStringsTable);
            final SheetIterator sheetsData = (SheetIterator)xssfReader.getSheetsData();
            m_countingSheetStream = getSheetStreamWithSheetName(sheetsData, getSelectedSheet());
            m_sheetSize = m_countingSheetStream.available();

            // create the parser
            return new XLSXParserRunnable(this, m_config, xmlReader, xssfReader, sharedStringsTable,
                use1904Windowing(xssfReader));
        } catch (SAXException | OpenXML4JException | ParserConfigurationException e) {
            throw new IllegalStateException(e.getMessage(), e);
        }
    }

    private static CountingInputStream getSheetStreamWithSheetName(final SheetIterator sheetsData,
        final String selectedSheet) throws IOException {
        while (sheetsData.hasNext()) {
            @SuppressWarnings("resource") // manually closed
            final InputStream sheetIterator = sheetsData.next();
            try {
                if (sheetsData.getSheetName().equals(selectedSheet)) {
                    return new CountingInputStream(sheetIterator);
                } else {
                    sheetIterator.close();
                }
            } catch (Exception e) { // NOSONAR catch anything to make sure the stream is closed
                sheetIterator.close();
            }
        }
        // should never happen as we checked for it already
        throw new IllegalStateException("No sheet with name '" + selectedSheet + "' found.");
    }

    private static boolean use1904Windowing(final XSSFReader xssfReader) {
        try (final InputStream workbookXml = xssfReader.getWorkbookData()) {
            final WorkbookDocument doc = WorkbookDocument.Factory.parse(workbookXml);
            final CTWorkbook wb = doc.getWorkbook();
            final CTWorkbookPr prefix = wb.getWorkbookPr();
            return prefix.getDate1904();
        } catch (Exception e) { // NOSONAR
            // if anything goes wrong, we just assume false
            return false;
        }
    }

    @Override
    public Map<String, Boolean> getSheetNames() {
        return m_sheetNames;
    }

    @Override
    public Set<Integer> getHiddenColumns() {
        return m_hiddenColumns;
    }

    @Override
    public void closeResources() throws IOException {
        if (m_countingSheetStream != null) {
            m_countingSheetStream.close();
        }
        if (m_opc != null) {
            m_opc.close();
        }
    }

    @Override
    public OptionalLong getMaxProgress() {
        return m_sheetSize < 0 ? OptionalLong.empty() : OptionalLong.of(m_sheetSize);
    }

    @Override
    public long getProgress() {
        return m_countingSheetStream.getByteCount();
    }

    private class XLSXParserRunnable extends ExcelParserRunnable {

        private final XSSFReader m_xssfReader;

        private final ReadOnlySharedStringsTable m_sharedStringsTable;

        private final KNIMEDataFormatter m_dataFormatter;

        private final XMLReader m_xmlReader;

        XLSXParserRunnable(final ExcelRead read, final TableReadConfig<ExcelTableReaderConfig> config,
            final XMLReader xmlReader, final XSSFReader xssfReader, final ReadOnlySharedStringsTable sharedStringsTable,
            final boolean use1904Windowing) {
            super(read, config);
            m_xmlReader = xmlReader;
            m_xssfReader = xssfReader;
            m_sharedStringsTable = sharedStringsTable;
            m_dataFormatter = new KNIMEDataFormatter(use1904Windowing, m_use15DigitsPrecision);
        }

        @Override
        protected void parse() throws Exception {
            final ExcelTableReaderSheetContentsHandler sheetContentsHandler =
                new ExcelTableReaderSheetContentsHandler();
            m_xmlReader.setContentHandler(new KNIMEXSSFSheetXMLHandler(m_xssfReader.getStylesTable(),
                m_sharedStringsTable, sheetContentsHandler, m_dataFormatter, false));
            m_xmlReader.parse(new InputSource(m_countingSheetStream));
        }

        class ExcelTableReaderSheetContentsHandler extends AbstractKNIMESheetContentsHandler {

            private int m_currentRowIdx = -1;

            private int m_lastNonEmptyRowIdx = -1;

            private int m_currentCol = -1;

            private boolean m_currentRowIsHiddenAndSkipped;

            private final List<ExcelCell> m_row = new ArrayList<>();

            private ExcelCell m_rowId;

            @Override
            public void startRow(final int rowIdx) {
                m_currentRowIsHiddenAndSkipped = m_skipHiddenRows && isHiddenRow();
                m_currentRowIdx = rowIdx;
                m_currentCol = -1;
            }

            @Override
            public void endRow(final int rowIdx) {
                if (!isRowEmpty(m_row)) {
                    // if there were empty rows in between two non-empty rows, output these (we need to make sure that
                    // there is a non-empty row after an empty row)
                    outputEmptyRows(m_currentRowIdx - m_lastNonEmptyRowIdx - 1);
                    m_lastNonEmptyRowIdx = m_currentRowIdx;
                    m_hiddenColumns = getHiddenCols();
                    // insert the row id at the beginning
                    insertRowIDAtBeginning(m_row, m_rowId);
                    // add the non-empty row the the queue
                    addToQueue(m_row, m_currentRowIsHiddenAndSkipped);
                }
                m_rowId = null;
                m_row.clear();
            }

            @Override
            public void cell(String cellReference, final String formattedValue, final XSSFComment comment) {
                m_currentCol++;
                if (cellReference == null) {
                    // gracefully handle missing CellRef here in a similar way as XSSFCell does.
                    // according to the API description the cellReference argument shouldn't be null,
                    // though it can be for files created by third party software (see e.g. AP-9380)
                    cellReference = new CellAddress(m_currentRowIdx, m_currentCol).formatAsString();
                }

                final int currentColCount = m_currentCol;
                m_currentCol = new CellReference(cellReference).getCol();

                final Optional<ExcelCell> numericCell = m_dataFormatter.getAndResetExcelCell();
                final ExcelCell excelCell = numericCell.orElseGet(() -> createExcelCellFromStringValue(formattedValue));
                if (!isColHiddenAndSkipped(m_currentCol, getHiddenCols())) {
                    if (isColRowID(m_currentCol)) {
                        // check if cells were missing and insert nulls if so
                        handleMissingCells(cellReference, currentColCount);
                        m_rowId = excelCell;
                    } else if (isColIncluded(m_currentCol)) {
                        // check if cells were missing and insert nulls if so
                        handleMissingCells(cellReference, currentColCount);
                        m_row.add(excelCell);
                    }
                } else if (isColRowID(m_currentCol)) {
                    m_rowId = excelCell;
                }
            }

            private ExcelCell createExcelCellFromStringValue(final String formattedValue) {
                final ExcelCell excelCell;
                final KNIMEXSSFDataType cellType = getCellType();
                switch (cellType) {
                    case BOOLEAN:
                        excelCell = new ExcelCell(KNIMECellType.BOOLEAN, formattedValue.equals("TRUE"));
                        break;
                    case ERROR:
                        excelCell = ExcelCellUtils.createErrorCell(m_config);
                        break;
                    case FORMULA:
                        excelCell = new ExcelCell(KNIMECellType.STRING, formattedValue);
                        break;
                    case STRING:
                        excelCell = new ExcelCell(KNIMECellType.STRING, formattedValue);
                        break;
                    case NUMBER_OR_DATE:
                    default:
                        throw new IllegalStateException("Unexpected data type: " + cellType);
                }
                return excelCell;
            }

            @Override
            public void headerFooter(final String text, final boolean isHeader, final String tagName) {
                // do nothing, we do not treat header and footer
            }

            private void handleMissingCells(final String cellReference, final int currentColCount) {
                final int thisCol = new CellReference(cellReference).getCol();
                final int missedCols = thisCol - currentColCount;
                for (int currentCol = currentColCount; currentCol < currentColCount + missedCols; currentCol++) {
                    if (isColIncluded(currentCol) && !isColHiddenAndSkipped(currentCol, getHiddenCols())
                        && !isColRowID(currentCol)) {
                        m_row.add(null);
                    }
                }
            }

            private boolean isColHiddenAndSkipped(final int col, final Set<Integer> hiddenCols) {
                return m_skipHiddenCols && hiddenCols.contains(col);
            }
        }
    }

}
