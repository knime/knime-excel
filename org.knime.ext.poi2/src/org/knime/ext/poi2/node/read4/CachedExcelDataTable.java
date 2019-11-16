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
 *   2 Sep 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read4;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.nio.file.Path;
import java.text.DateFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeParseException;
import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Map.Entry;
import java.util.NoSuchElementException;
import java.util.OptionalDouble;
import java.util.OptionalLong;
import java.util.Set;
import java.util.SortedMap;
import java.util.TimeZone;
import java.util.TreeMap;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import org.apache.commons.io.input.CountingInputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.data.BooleanValue;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTable;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataTableSpecCreator;
import org.knime.core.data.DataType;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.IntValue;
import org.knime.core.data.MissingCell;
import org.knime.core.data.MissingValue;
import org.knime.core.data.RowIterator;
import org.knime.core.data.RowKey;
import org.knime.core.data.StringValue;
import org.knime.core.data.def.BooleanCell;
import org.knime.core.data.def.BooleanCell.BooleanCellFactory;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.data.time.localdatetime.LocalDateTimeCellFactory;
import org.knime.core.data.time.localdatetime.LocalDateTimeValue;
import org.knime.core.data.util.CancellableReportingInputStream;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.util.FileUtil;
import org.knime.core.util.Pair;
import org.knime.core.util.ThreadUtils;
import org.knime.ext.poi2.node.read4.KNIMEXSSFSheetXMLHandler.KNIMESheetContentsHandler;
import org.knime.ext.poi2.node.read4.POIUtils.ColumnTypeCombinator;
import org.knime.ext.poi2.node.read4.POIUtils.StopProcessing;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import com.google.common.base.Strings;
import com.google.common.collect.Interner;
import com.google.common.collect.Interners;

/**
 * Caches the read data and creates iterators based on that.
 *
 * @author Gabor Bakos
 */
final class CachedExcelTable {

    private static final AtomicInteger CACHED_THREAD_POOL_INDEX = new AtomicInteger();

    /**
     * Threadpool that can be used to create new cached tables.
     */
    private static final ExecutorService CACHED_THREAD_POOL = Executors.newCachedThreadPool(
        r -> ThreadUtils.threadWithContext(r, "KNIME-XLS-Parser-" + CACHED_THREAD_POOL_INDEX.getAndIncrement()));

    /**
     * Container for a cached value.
     */
    static final class Content {
        private final String m_originalValue;
        private final String m_valueAsString;
        private final ActualDataType m_type;

        /**
         * @param valueAsString The contained value in a parseable {@link String}.
         * @param type The {@link ActualDataType} of the content.
         */
        Content(final String valueAsString, final ActualDataType type) {
            m_originalValue = valueAsString;
            m_valueAsString = valueAsString;
            m_type = type;
        }

        /**
         * @param valueAsString The contained value in a parseable {@link String}.
         * @param originalValue The original {@link String} format of the value (for example for dates or numbers the
         *            formatted value).
         * @param type The {@link ActualDataType} of the content.
         */
        Content(final String valueAsString, final String originalValue, final ActualDataType type) {
            m_originalValue = originalValue;
            m_valueAsString = valueAsString;
            m_type = type;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public String toString() {
            return String.format("Content [valueAsString=%s, type=%s]", m_valueAsString, m_type);
        }

    }

    @SuppressWarnings("unchecked")
    private final ArrayList<Content>[] m_contents = new ArrayList[16_384];

    private final Set<Integer> m_hiddenColumns = new HashSet<>();

    private final DateFormat m_dateFormat = new SimpleDateFormat("yyyy-MM-dd"),
            m_dateAndTimeFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");

    {
        m_dateFormat.setTimeZone(TimeZone.getTimeZone(ZoneOffset.UTC));
        m_dateAndTimeFormat.setTimeZone(TimeZone.getTimeZone(ZoneOffset.UTC));
    }

    private volatile boolean m_incomplete = true;

    private int m_lastColumnIndex = Integer.MIN_VALUE;
    private int m_lastRowIndex = Integer.MIN_VALUE;

    private final Map<Integer, Content> m_rowMap = new HashMap<>();

    /** Hide constructor. */
    private CachedExcelTable() {
    }

    /**
     * Visits the sheet for xlsx streaming and fills {@link CachedExcelTable}.
     */
    final class KNIMESheetContentVisitor extends KNIMESheetContentsHandler.Abstract {
        private int m_currentRow = -1;

        private KNIMEDataFormatter m_formatter;

        private Map<Integer, Content> m_currentRowMap = new HashMap<>();

        private ExecutionMonitor m_exec;

        private final Interner<String> m_stringInterner = Interners.newWeakInterner();

        private final Supplier<OptionalDouble> m_progressSupplier;

        //private final ArrayBlockingQueue<Optional<DataRow>> m_readRows;

        //private volatile boolean m_finished = false;
        /**
         * @param formatter
         * @param exec
         * @param progressSupplier a supplier that reports in [0, 1], or null if unknown progress
         */
        public KNIMESheetContentVisitor(final KNIMEDataFormatter formatter, final ExecutionMonitor exec,
            final Supplier<OptionalDouble> progressSupplier) {
            m_formatter = formatter;
            m_exec = exec;
            m_progressSupplier = progressSupplier;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void startRow(final int rowNum) {
            super.startRow(rowNum);
            checkCancelled();
            m_currentRow = rowNum;
            m_currentRowMap.clear();
            if (m_exec != null) {
                // fallback: we do not know the actual size of the XLS, provide a logarithmic progress for max row count
                double progress = m_progressSupplier.get().orElseGet(() -> Math.log1p(rowNum) / Math.log(1_048_576d));
                m_exec.setProgress(progress, () -> "Reading row " + rowNum);
            }
        }

        /**
         * Checks whether execution has been cancelled trhough the monitor or the current thread is intterrupted.
         */
        private void checkCancelled() {
            if (m_exec != null) {
                try {
                    m_exec.checkCanceled();
                } catch (CanceledExecutionException e) {
                    throw new POIUtils.StopProcessing();
                }
            }
            final Thread currentThread = Thread.currentThread();
            if (currentThread.isInterrupted()) {
                currentThread.interrupt();
                throw new POIUtils.StopProcessing();
            }
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void endRow(final int rowNum) {
            super.endRow(rowNum);
            checkCancelled();
            for (Entry<Integer, Content> entry : m_currentRowMap.entrySet()) {
                int col = entry.getKey().intValue();
                if (m_contents[col] == null) {
                    m_contents[col] = new ArrayList<>(16);
                }
                ArrayList<Content> column = m_contents[col];
                while (column.size() < rowNum) {
                    column.add(null);
                }
                column.add(entry.getValue());
            }
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void cell(String cellReference, final String formattedValue, final XSSFComment comment) {
            checkCancelled();
            if (cellReference == null) {
                // gracefully handle missing CellRef here in a similar way as XSSFCell does;
                // according to the API description the cellReference argument shouldn't be null,
                // though it can be as seen in the file attached to AP-9380
                cellReference = new CellAddress(m_currentRow, 0).formatAsString();
            }

            // Did we miss any cells?
            CellReference cellReferenceObj = new CellReference(cellReference);
            int thisCol = cellReferenceObj.getCol();
            //int thisRow = cellReferenceObj.getRow();
            final ActualDataType type;
            final String valueAsString;
            switch (m_dataType) {
                case NUMBER_OR_DATE:
                    if (formattedValue == null) {
                        //Most probably a missing cell with a comment
                        valueAsString = null;
                        type = ActualDataType.MISSING;
                    } else /* Number or string?*/  if (formattedValue.startsWith(KNIMEDataFormatter.DATE_PREFIX)) {
                        type = ActualDataType.DATE;
                        valueAsString = formattedValue.substring(KNIMEDataFormatter.DATE_PREFIX.length());
                    } else if (formattedValue.startsWith(KNIMEDataFormatter.NUMBER_PREFIX)) {
                        type = POIUtils.isInteger(formattedValue.substring(KNIMEDataFormatter.NUMBER_PREFIX.length()))
                            ? ActualDataType.NUMBER_INT : ActualDataType.NUMBER_DOUBLE;
                        valueAsString = formattedValue.substring(KNIMEDataFormatter.NUMBER_PREFIX.length());
                    } else if (formattedValue.startsWith(KNIMEDataFormatter.BOOL_PREFIX)) {
                        type = ActualDataType.BOOLEAN;
                        valueAsString = toBooleanString(formattedValue);
                    } else {//Probably should never happen as we not include them in fill... methods.
                        type = ActualDataType.MISSING;
                        valueAsString = null;
                    }
                    break;
                case STRING:
                    type = ActualDataType.STRING;
                    valueAsString = formattedValue;
                    break;
                case BOOLEAN:
                    type = ActualDataType.BOOLEAN;
                    valueAsString = formattedValue.endsWith("+")
                        ? formattedValue.substring(0, formattedValue.length() - 1) : formattedValue;
                    break;
                case ERROR:
                    type = ActualDataType.ERROR;
                    valueAsString = formattedValue;
                    break;
                case FORMULA:
                    //TODO should we visit again to find positive and negative infinity formulae?
                    // Number or string?
                    if (formattedValue == null) {
                        type = ActualDataType.MISSING_FORMULA;
                        valueAsString = null;
                    } else if (formattedValue.startsWith(KNIMEDataFormatter.DATE_PREFIX)) {
                        type = ActualDataType.DATE_FORMULA;
                        valueAsString = formattedValue.substring(KNIMEDataFormatter.DATE_PREFIX.length());
                    } else if (formattedValue.startsWith(KNIMEDataFormatter.NUMBER_PREFIX)) {
                        type = POIUtils.isInteger(formattedValue.substring(1)) ? ActualDataType.NUMBER_INT_FORMULA
                            : ActualDataType.NUMBER_DOUBLE_FORMULA;
                        valueAsString = formattedValue.substring(KNIMEDataFormatter.NUMBER_PREFIX.length());
                    } else if (formattedValue.startsWith(KNIMEDataFormatter.BOOL_PREFIX)) {
                        type = ActualDataType.BOOLEAN_FORMULA;
                        valueAsString = toBooleanString(formattedValue);
                    } else {
                        type = ActualDataType.STRING_FORMULA;
                        valueAsString = formattedValue;
                    }
                    break;
                default:
                    throw new IllegalStateException("Unknown type: " + m_dataType);
            }
            //Safe call as only this thread uses it
            String lastOriginalFormattedValue = m_formatter.getLastOriginalFormattedValue();
            if (ActualDataType.isBoolean(type)) {
                if (StringUtils.startsWith(lastOriginalFormattedValue, KNIMEDataFormatter.BOOL_PREFIX)) {
                    lastOriginalFormattedValue =
                        lastOriginalFormattedValue.substring(KNIMEDataFormatter.BOOL_PREFIX.length());
                } else {
                    lastOriginalFormattedValue = m_stringInterner.intern(valueAsString);
                }
                m_currentRowMap.put(thisCol,
                    new Content(m_stringInterner.intern(valueAsString), lastOriginalFormattedValue, type));
            } else if (ActualDataType.isDate(type) || ActualDataType.isNumber(type)) {
                if (StringUtils.startsWith(lastOriginalFormattedValue, KNIMEDataFormatter.DATE_PREFIX)) {
                    lastOriginalFormattedValue =
                        lastOriginalFormattedValue.substring(KNIMEDataFormatter.DATE_PREFIX.length());
                } else if (StringUtils.startsWith(lastOriginalFormattedValue, KNIMEDataFormatter.NUMBER_PREFIX)) {
                    lastOriginalFormattedValue =
                        lastOriginalFormattedValue.substring(KNIMEDataFormatter.NUMBER_PREFIX.length());
                }
                m_currentRowMap.put(thisCol,
                    new Content(m_stringInterner.intern(valueAsString), lastOriginalFormattedValue, type));
            } else if (valueAsString != null) {
                m_currentRowMap.put(thisCol, new Content(m_stringInterner.intern(valueAsString), type));
            }
        }

        /**
         * @param formattedValue
         * @return
         */
        private String toBooleanString(final String formattedValue) {
            final String tmp = formattedValue.substring(KNIMEDataFormatter.BOOL_PREFIX.length()).toLowerCase();
            return tmp.endsWith("+") ? tmp.substring(0, tmp.length() - 1) : tmp;
        }
    }

    /**
     *
     * Constructs {@link CachedExcelTable} using the xlsx streaming representation.
     *
     * @param stream The workbook's stream (to handle knime:// urls too, we do not create it on the thread of
     *            threadpool).
     * @param sheet The sheet's name.
     * @param locale The {@link Locale} to use.
     * @param exec The {@link ExecutionMonitor} to use.
     * @param incompleteResult The container for the incomplete result, can be {@code null}.
     * @param reevaluate Should we reevaluate the formulae?
     * @return The {@link Future} representing the computation of {@link CachedExcelTable}.
     */
    static Future<CachedExcelTable> fillCacheFromXlsxStreaming(final Path path, final InputStream stream,
        final String sheet, final Locale locale, final ExecutionMonitor exec,
        final AtomicReference<CachedExcelTable> incompleteResult) {
        //Probably could be improved to let visit the intermediate results while generating, but it is complicated.
        return CACHED_THREAD_POOL.submit(ThreadUtils.callableWithContext(() -> {
            LocaleUtil.setUserLocale(locale);
            final CachedExcelTable table = new CachedExcelTable();
            final KNIMEDataFormatter formatter = new KNIMEDataFormatter(locale);

            try (final OPCPackage opc = OPCPackage.open(stream)) {
                final XSSFReader xssfReader = new XSSFReader(opc);
                final ReadOnlySharedStringsTable readOnlySharedStringsTable = new ReadOnlySharedStringsTable(opc, false);
                boolean sheetAvailable = false;
                for (final SheetIterator sheetIt = (SheetIterator)xssfReader.getSheetsData(); sheetIt.hasNext();) {
                    InputStream is = sheetIt.next();
                    if (sheet.equals(sheetIt.getSheetName())) { // not closed here; method arg to be closed by caller
                        sheetAvailable = true;
                        final Supplier<OptionalDouble> progressSupplier;
                        long sheetSize;
                        if (is instanceof ByteArrayInputStream) { // debugger told me this is often a BAIS
                            sheetSize = ((ByteArrayInputStream)is).available();
                        } else {
                            sheetSize = sheetIt.getSheetPart().getSize();
                        }
                        if (sheetSize >= 0L) {
                            final double asDouble = sheetSize;
                            @SuppressWarnings("resource")
                            CountingInputStream countingStream = new CountingInputStream(is);
                            is = countingStream;
                            progressSupplier = () -> OptionalDouble.of(countingStream.getByteCount() / asDouble);
                        } else {
                            progressSupplier = () -> OptionalDouble.empty();
                        }
                        final InputSource sheetSource = new InputSource(is);
                        final XMLReader sheetParser = SAXHelper.newXMLReader();
                        final KNIMESheetContentVisitor sheetContentsHandler =
                            table.new KNIMESheetContentVisitor(formatter, exec, progressSupplier);

                        final KNIMEXSSFSheetXMLHandler handler = new KNIMEXSSFSheetXMLHandler(
                            xssfReader.getStylesTable(), /*could be null*/sheetIt.getSheetComments(),
                            readOnlySharedStringsTable, sheetContentsHandler, formatter, false);
                        sheetParser.setContentHandler(handler);
                        try {
                            sheetParser.parse(sheetSource);
                            exec.setProgress(1.0, () -> "Reading finished");
                            table.m_incomplete = false;
                        } catch (RuntimeException e) {//Includes StopProcessing
                            throw e;
                        } catch (Exception e) {
                            throw new RuntimeException(e);
                        } finally {
                            table.m_hiddenColumns.addAll(handler.getHiddenColumns());
                        }
                        break;
                    }
                }
                if (!sheetAvailable) {
                    throw new IOException("Workbook \"" + path.getFileName().toString()
                        + "\" does not contain a sheet called \"" + sheet + "\"");
                }
            } catch (StopProcessing e) {
                if (incompleteResult != null) {
                    incompleteResult.set(table);
                }
                throw e;
            }
            return table;
        }));
    }

    /**
     * Constructs {@link CachedExcelTable} using the DOM-based representation.
     *
     * @param path The path of the workbook.
     * @param stream The workbook's stream (to handle knime:// urls too, we do not create it on the thread of
     *            threadpool).
     * @param sheet The sheet's name.
     * @param locale The {@link Locale} to use.
     * @param reevaluate Should we reevaluate the formulae?
     * @param exec The {@link ExecutionMonitor} to use.
     * @param incompleteResult The container for the incomplete result, can be {@code null}.
     * @return The {@link Future} representing the computation of {@link CachedExcelTable}.
     */
    static Future<CachedExcelTable> fillCacheFromDOM(final Path path, final InputStream stream,
        final String sheet, final Locale locale, final boolean reevaluate, final ExecutionMonitor exec,
        final AtomicReference<CachedExcelTable> incompleteResult) {
        return CACHED_THREAD_POOL.submit(ThreadUtils.callableWithContext(() -> {
            LocaleUtil.setUserLocale(locale);
            OptionalLong fileSize = getFileSize(path.toString());
            CachedExcelTable table = new CachedExcelTable();
            ExecutionMonitor workbookCreateProgress = exec.createSubProgress(.2);
            ExecutionMonitor workBookParseProgress = exec.createSubProgress(.8);
            exec.setMessage("Reading workbooks...");
            try (final CancellableReportingInputStream cancellableStream =
                new CancellableReportingInputStream(stream, workbookCreateProgress, fileSize.orElse(-1L));
                    final Workbook workbook = WorkbookFactory.create(cancellableStream)) {
                workbookCreateProgress.setProgress(1.0);
                exec.setMessage("Parsing workbooks...");
                final Sheet sheetXls = workbook.getSheet(sheet);
                if (sheetXls == null) {
                    throw new IOException("Workbook \"" + path.getFileName().toString()
                        + "\" does not contain a sheet called \"" + sheet + "\"");
                }
                final FormulaEvaluator evaluator =
                    reevaluate ? workbook.getCreationHelper().createFormulaEvaluator() : null;
                Thread currentThread = Thread.currentThread();
                boolean date1904;
                if (workbook instanceof XSSFWorkbook) {
                    @SuppressWarnings("resource")
                    final XSSFWorkbook xssfWorkbook = (XSSFWorkbook)workbook;
                    date1904 = xssfWorkbook.isDate1904();
                } else if (workbook instanceof HSSFWorkbook) {
                    @SuppressWarnings("resource")
                    final HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                    date1904 = hssfWorkbook.getInternalWorkbook().isUsing1904DateWindowing();
                } else {
                    //Probably unsupported
                    date1904 = false;
                }

                for (Row row : sheetXls) {
                    workBookParseProgress.setProgress(((double)row.getRowNum()) / sheetXls.getLastRowNum(),
                        () -> "Row: " + row.getRowNum());
                    if (currentThread.isInterrupted()) {
                        currentThread.interrupt();
                        throw new StopProcessing();
                    }
                    exec.checkCanceled();
                    Map<Integer, Content> rowMap = new HashMap<>();
                    for (Cell cell : row) {
                        rowMap.put(cell.getColumnIndex(), table.createContentFromXLCell(cell, evaluator, date1904));
                    }
                    if (!rowMap.values().stream().allMatch(c -> ActualDataType.isMissing(c.m_type))) {
                        for (Entry<Integer, Content> entry : rowMap.entrySet()) {
                            int col = entry.getKey().intValue();
                            if (table.m_contents[col] == null) {
                                table.m_contents[col] = new ArrayList<>(16);
                            }
                            ArrayList<Content> column = table.m_contents[col];
                            while (column.size() < row.getRowNum()) {
                                column.add(null);
                            }
                            column.add(entry.getValue());
                        }
                    }
                }
                workBookParseProgress.setProgress(1.0);
                int lastCol = lastNonNull(table.m_contents);
                for (int i = 1; i <= lastCol + 1; ++i) {
                    if (sheetXls.isColumnHidden(i - 1)) {
                        table.m_hiddenColumns.add(i);
                    }
                }
                table.m_incomplete = false;
            } catch (StopProcessing | CancellationException e) {
                if (incompleteResult != null) {
                    incompleteResult.set(table);
                }
                throw e;
            }
            return table;
        }));
    }

    /**
     * @param contents
     * @return
     */
    private static int lastNonNull(final Object[] contents) {
        int lastNonNull = -1;
        for (int i = 0; i < contents.length; i++) {
            if (contents[i] != null) {
                lastNonNull = i;
            }
        }
        return lastNonNull;
    }

    /** The size of the file as per argument or an empty {@link OptionalLong} if not possible
     * (e.g. remote file or invalid file).
     * @param loc The location
     * @return The size of the file.
     */
    private static OptionalLong getFileSize(final String loc) {
        if (Strings.isNullOrEmpty(loc)) {
            return OptionalLong.empty();
        }
        try {
            File f = FileUtil.getFileFromURL(FileUtil.toURL(loc));
            if (f != null) {
                return OptionalLong.of(f.length());
            }
        } catch (MalformedURLException | IllegalArgumentException e) {
        }
        return OptionalLong.empty();
    }

    /**
     * Used for DOM-based reading. Converts {@code cell} to {@link Content}.
     *
     * @param cell A {@link Cell}.
     * @param evaluator The evaluator.
     * @param use1904windowing
     */
    private Content createContentFromXLCell(final Cell cell, final FormulaEvaluator evaluator, final boolean use1904windowing) {
        if (cell == null) {
            return new Content(null, ActualDataType.MISSING);
        }
        // determine the type
        return createContentFromXLCell(cell, evaluator, cell.getCellType(), use1904windowing);

    }

    /**
     * Used for DOM-based reading. Converts {@code cell} to {@link Content}.
     *
     * @param cell A {@link Cell}.
     * @param evaluator The evaluator.
     * @param cellType The type of {@code cell} (in case it is not the same as {@code cell}'s {@link Cell#getCellType()}
     *            ).
     * @param use1904windowing
     */
    private Content createContentFromXLCell(final Cell cell, final FormulaEvaluator evaluator, final int cellType, final boolean use1904windowing) {
        final int colIdx = cell.getColumnIndex();
        DataFormatter formatter = new DataFormatter();
        Format cellFormat;
        try {
            cellFormat = formatter.createFormat(cell);
        } catch (RuntimeException e) {
            cellFormat = null;
        }
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK:
                return new Content(null, ActualDataType.MISSING);
            case Cell.CELL_TYPE_BOOLEAN:
                boolean b = cell.getBooleanCellValue();
                return new Content(Boolean.toString(b),
                    cellFormat == null ? Boolean.toString(b) : cellFormat.format(Double.valueOf(b ? 1d : 0d)),
                    ActualDataType.BOOLEAN);
            case Cell.CELL_TYPE_ERROR:
                return new Content(Integer.toString(cell.getErrorCellValue()), ActualDataType.ERROR);
            case Cell.CELL_TYPE_FORMULA:
                CellValue cellValue = null;
                // Check if the formula is the replacement formula for infinity (see XLSWriter2.write())
                if (cell.getCellFormula().equals("1/0")) {
                    return new Content(Double.toString(Double.POSITIVE_INFINITY),
                        Double.toString(Double.POSITIVE_INFINITY), ActualDataType.NUMBER_DOUBLE_FORMULA);
                } else if (cell.getCellFormula().equals("-1/0")) {
                    return new Content(Double.toString(Double.NEGATIVE_INFINITY),
                        Double.toString(Double.NEGATIVE_INFINITY), ActualDataType.NUMBER_DOUBLE_FORMULA);
                }
                if (evaluator == null) {
                    CheckUtils.checkState(cell.getCachedFormulaResultType() != Cell.CELL_TYPE_FORMULA,
                        "Formula cannot create another formula");
                    //Use cached value.
                    return createContentFromXLCell(cell, evaluator, cell.getCachedFormulaResultType(), use1904windowing);
                }
                try {
                    cellValue = evaluator.evaluate(cell);
                } catch (Exception e) {
                    return new Content(e.getMessage(), ActualDataType.ERROR_FORMULA);
                }
                switch (cellValue.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        boolean bvalue = cellValue.getBooleanValue();
                        return new Content(Boolean.toString(bvalue),
                            cellFormat == null ? Boolean.toString(bvalue) : cellFormat.format(bvalue ? 1d : 0d),
                            ActualDataType.BOOLEAN_FORMULA);
                    case Cell.CELL_TYPE_NUMERIC:
                        //Checking for data, based on  DateUtil#isCellDateFormatted(Cell)
                        double d = cellValue.getNumberValue();
                        if (DateUtil.isADateFormat(cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString())) {
                            //Date date = cell.getDateCellValue();
                            Date date = DateUtil.getJavaDate(cellValue.getNumberValue(), use1904windowing);
                            Date dateUTC = DateUtil.getJavaDate(d, use1904windowing, TimeZone.getTimeZone(ZoneOffset.UTC), true);
                            return new Content(
                                d == (long)d ? m_dateFormat.format(dateUTC) : m_dateAndTimeFormat.format(dateUTC),
                                cellFormat == null ? m_dateFormat.format(date) : cellFormat.format(date),
                                ActualDataType.DATE_FORMULA);
                        } else {
                            double value = cellValue.getNumberValue();
                            if (POIUtils.isInteger(Double.toString(value))) {
                                return new Content(Integer.toString((int)value),
                                    cellFormat == null ? Integer.toString((int)value) : cellFormat.format(value),
                                    ActualDataType.NUMBER_INT_FORMULA);
                            } else {
                                return new Content(/*KNIMEDataFormatter.NUMBER_PREFIX + */Double.toString(value),
                                    cellFormat == null ? Double.toString(value) : cellFormat.format(value),
                                    ActualDataType.NUMBER_DOUBLE_FORMULA);
                            }
                        }
                    case Cell.CELL_TYPE_STRING:
                        return new Content(cellValue.getStringValue(), ActualDataType.STRING_FORMULA);
                    case Cell.CELL_TYPE_BLANK:
                        return new Content(null, ActualDataType.MISSING_FORMULA);
                    case Cell.CELL_TYPE_ERROR:
                        return new Content("Err (#" + (cellValue.getErrorValue() & 0xff) + ")",
                            ActualDataType.ERROR_FORMULA);
                    case Cell.CELL_TYPE_FORMULA:
                        throw new IllegalStateException("Invalid formula result type in column idx " + colIdx
                        /*+ ", sheet '" + settings.getSheetName()*/ + "', row " + cell.getRowIndex());
                }
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    final Date date = cell.getDateCellValue();
                    final double d = cell.getNumericCellValue();
                    Date dateUTC = DateUtil.getJavaDate(d, use1904windowing, TimeZone.getTimeZone(ZoneOffset.UTC), true);
                    return new Content(
                        d == (long)d ? m_dateFormat.format(dateUTC) : m_dateAndTimeFormat.format(dateUTC),
                        cellFormat == null ? m_dateAndTimeFormat.format(date) : cellFormat.format(date),
                        ActualDataType.DATE);
                } else {
                    double value = cell.getNumericCellValue();
                    if (POIUtils.isInteger(Double.toString(value))) {
                        return new Content(Integer.toString((int)value),
                            cellFormat == null ? Integer.toString((int)value) : cellFormat.format(value),
                            ActualDataType.NUMBER_INT);
                    } else {
                        return new Content(/*KNIMEDataFormatter.NUMBER_PREFIX + */Double.toString(value),
                            cellFormat == null ? Double.toString(value) : cellFormat.format(value),
                            ActualDataType.NUMBER_DOUBLE);
                    }
                }
            case Cell.CELL_TYPE_STRING:
                return new Content(cell.getRichStringCellValue().getString(), ActualDataType.STRING);
            default:
                throw new IllegalStateException("Invalid cell type in column idx " + colIdx + ", sheet '"
                /*+ settings.getSheetName() */ + "', row " + cell.getRowIndex());
        }
    }

    /**
     * Generates a {@link DataTable} based on {@code rawSettings}.
     *
     * @param rawSettings The settings to be used to generate the table.
     * @param resultExcelToKNIME The mapping of excel {@code 1-based} columns to KNIME output columns {@code 0}-based.
     *            This will be a result.
     * @param uniquifier the {@link ValueUniquifier} to use accross all files
     * @return The {@link DataTable}.
     * @throws InvalidSettingsException Some settings in {@code rawSettings} is incorrect.
     */
    DataTable createDataTable(final Path path, final XLSUserSettings rawSettings, final Map<Integer, Integer> resultExcelToKNIME,
        final long rowNoToStart, final long totalNoOfPreviousRows, final ValueUniquifier uniquifier)
        throws InvalidSettingsException {
        final XLSUserSettings settings = XLSUserSettings.normalizeSettings(rawSettings);
        final Set<Integer> skippedCols = new HashSet<>();
        final DataTableSpec spec = createSpec(path, settings, skippedCols);
        final DataTableSpec knimeSpec = updatedSpec(spec);
        final Map<Integer, Integer> mapFromExcelColumnIndicesToKNIME = new HashMap<>();
        final DataTable dataTable = new DataTable() {

            {
                int exCol = 0, knCol = 0, allKnCols = spec.getNumColumns();
                while (knCol < allKnCols) {
                    if (!skippedCols.contains(exCol)) {
                        mapFromExcelColumnIndicesToKNIME.put(exCol, knCol++);
                    }
                    ++exCol;
                }
            }

            @Override
            public DataTableSpec getDataTableSpec() {
                return knimeSpec;
            }

            @Override
            public RowIterator iterator() {
                return new RowIterator() {
                    private int m_rowNumber = Math.max(0, settings.getFirstRow0());

                    private Entry<Integer, Map<Integer, Content>> m_nextRow;

                    private final long m_noOfRowsPreviousTables = totalNoOfPreviousRows;

                    private long m_lastRow = Math.max(0, settings.getFirstRow0()) - 1;
                    private long m_rowKeyIndex = rowNoToStart;

                    private boolean m_calledNext = true;

                    private boolean m_hasNext;

                    private DataCell[] m_emptyRow = new DataCell[spec.getNumColumns()],
                            m_cells = new DataCell[spec.getNumColumns()];
                    {
                        Arrays.fill(m_emptyRow, DataType.getMissingCell());
                    }

                    private RowKey m_rowKey;

                    @Override
                    public boolean hasNext() {
                        if (!m_calledNext) {
                            return m_hasNext;
                        }
                        Arrays.fill(m_cells, DataType.getMissingCell());
                        if (settings.getSkipEmptyRows()) {
                            Entry<Integer, Map<Integer, Content>> next;
                            do {
                                next = nextEntry();
                            } while (next != null && (isAllMissing(next) || isSkipRow(next.getKey())));
                            m_hasNext = next != null;
                            if (next != null) {
                                m_lastRow = next.getKey().longValue() + m_noOfRowsPreviousTables;
                                m_rowKeyIndex++;
                                process(next);
                            }
                        } else {//keep empty rows
                            if (m_nextRow == null) {
                                m_nextRow = nextEntry();
                            }
                            if (m_nextRow == null) {//We are after the last real row
                                m_lastRow++;
                                //TODO handle the case when it should be skipped.
                                m_hasNext = m_lastRow >= settings.getFirstRow0() && m_lastRow <= settings.getLastRow0() && m_lastRow < numOfRows();
                                m_rowKeyIndex++;
                                rowKey(new AbstractMap.SimpleImmutableEntry<>(Integer.valueOf((int)m_lastRow),
                                    Collections.emptyMap()));
                                //cells are already empty
                            } else if (m_lastRow + 1 == m_nextRow.getKey().longValue()) {
                                //next row should come unless it is to be skipped
                                while (m_nextRow != null && isSkipRow(m_nextRow.getKey())) {
                                    m_lastRow++;
                                    if (!isHeader(m_nextRow.getKey())) {
                                        m_rowKeyIndex++;
                                    }
                                    m_nextRow = nextEntry();
                                }
                                if (m_nextRow == null) {
                                    return hasNext();
                                }
                                m_lastRow++;
                                m_rowKeyIndex++;
                                process(m_nextRow);
                                m_hasNext = true;
                                m_nextRow = nextEntry();
                            } else {//empty row, but it might be before or after the first/last
                                m_lastRow++;
                                //TODO handle the case when it should be skipped.
                                m_hasNext = m_lastRow >= settings.getFirstRow0()
                                    && (settings.getLastRow0() < 0 || m_lastRow <= settings.getLastRow0());
                                rowKey(new AbstractMap.SimpleImmutableEntry<>(Integer.valueOf((int)m_lastRow),
                                    Collections.emptyMap()));
                                m_rowKeyIndex++;
                                //cells are already empty
                            }
                        }
                        m_calledNext = false;
                        return m_hasNext;
                    }

                    /**
                     * @param next The entry representing the row.
                     * @return If all non-skipped values in the row is missing, {@code true}, otherwise {@code false}.
                     */
                    private boolean isAllMissing(final Entry<Integer, Map<Integer, Content>> next) {
                        return next.getValue().entrySet().stream().allMatch(entry -> {
                            Integer key = entry.getKey();
                            Content content = entry.getValue();
                            Integer index = mapFromExcelColumnIndicesToKNIME.get(key);
                            return !mapFromExcelColumnIndicesToKNIME.containsKey(key)
                                || convertToCell(content.m_valueAsString, spec.getColumnSpec(index).getType())
                                    .isMissing();
                        });
                    }

                    /**
                     * Processes the row.
                     *
                     * @param next An entry representing a new row.
                     */
                    private void process(final Entry<Integer, Map<Integer, Content>> next) {
                        rowKey(next);
                        for (Entry<Integer, Content> cell : next.getValue().entrySet()) {
                            int col = cell.getKey();
                            final Content content = cell.getValue();
                            final String valueAsString = content.m_valueAsString;
                            final String original = content.m_originalValue;
                            final ActualDataType type = content.m_type;
                            //Necessary condition as there might be excessive missing value
                            if (mapFromExcelColumnIndicesToKNIME.containsKey(col)) {
                                final int idx = mapFromExcelColumnIndicesToKNIME.get(col);
                                DataType expectedType = spec.getColumnSpec(idx).getType();
                                if (valueAsString != null) {
                                    if (valueAsString.equals(settings.getMissValuePattern())) {
                                        expectedType = DataType.getMissingCell().getType();
                                    }
                                    switch (type) {
                                        case DATE:
                                        case DATE_FORMULA:
                                            m_cells[idx] = convertToCell(valueAsString, expectedType);
                                            break;
                                        case NUMBER:
                                        case NUMBER_DOUBLE:
                                        case NUMBER_INT:
                                        case NUMBER_FORMULA:
                                        case NUMBER_DOUBLE_FORMULA:
                                        case NUMBER_INT_FORMULA:
                                            if (expectedType.isCompatible(DoubleValue.class)) {
                                                m_cells[idx] = convertToCell(valueAsString, expectedType);
                                            } else {
                                                m_cells[idx] = convertToCell(original, expectedType);
                                            }
                                            break;
                                        case ERROR:
                                        case ERROR_FORMULA:
                                            if (settings.getUseErrorPattern()) {
                                                m_cells[idx] =
                                                    convertToCell(settings.getErrorPattern(), StringCell.TYPE);
                                            } else {
                                                m_cells[idx] =
                                                    convertToCell(valueAsString, DataType.getMissingCell().getType());
                                            }
                                            break;
                                        case BOOLEAN:
                                        case BOOLEAN_FORMULA:
                                        case MISSING:
                                        case MISSING_FORMULA:
                                            m_cells[idx] = convertToCell(valueAsString, expectedType);
                                            break;
                                        case STRING:
                                        case STRING_FORMULA:
                                            m_cells[idx] = convertToCell(valueAsString, expectedType);
                                            break;
                                        default:
                                            throw new UnsupportedOperationException("Not supported type: " + type);
                                    }
                                } else {
                                    m_cells[idx] = convertToCell(valueAsString, expectedType);
                                }
                            }
                        }
                    }

                    /**
                     * Sets the rowkey based on the row.
                     *
                     * @param next The next row raw representation.
                     */
                    private void rowKey(final Entry<Integer, Map<Integer, Content>> next) {
                        if (settings.getHasRowHeaders()) {
                            Content content = next.getValue().get(settings.getRowHdrCol0());
                            String rowKey =
                                content == null ? "" : content.m_valueAsString == null ? "" : content.m_valueAsString;
                            if (settings.getUniquifyRowIDs()) {
                                rowKey = uniquifier.uniquifyRowHeader(rowKey);
                            } else if (rowKey.isEmpty()) {
                                rowKey = RowKey.createRowKey(m_lastRow + m_noOfRowsPreviousTables).getString();
                            }
                            m_rowKey = new RowKey(rowKey);
                        } else if (settings.getKeepXLColNames()) {
                            m_rowKey = new RowKey(Long.toString(m_rowKeyIndex + 1L));
                        } else if (settings.isIndexContinuous()){
                            m_rowKey = RowKey.createRowKey(m_rowKeyIndex);
                        } else if (settings.isIndexSkipJumps()) {
                            m_rowKey = RowKey.createRowKey(m_lastRow + m_noOfRowsPreviousTables);
                        }
                    }

                    private DataCell convertToCell(final String valueAsString, final DataType type) {
                        if (valueAsString == null) {
                            return DataType.getMissingCell();
                        }
                        if (valueAsString.equals(settings.getMissValuePattern())) {
                            return new MissingCell(valueAsString);
                        }
                        try {
                            if (type.isCompatible(MissingValue.class)) {
                                return new MissingCell(valueAsString);
                            }
                            if (type.isCompatible(BooleanValue.class)) {
                                final String lowerCase = valueAsString.toLowerCase();
                                CheckUtils.checkState("true".equals(lowerCase) || "false".equals(lowerCase),
                                    "Not a boolean value: " + valueAsString);
                                return BooleanCellFactory.create(lowerCase);
                            }
                            if (type.isCompatible(IntValue.class)) {
                                try {
                                    double d = Double.parseDouble(valueAsString);
                                    if (d != (int)d) {
                                        //Just for the exception.
                                        Integer.parseInt(valueAsString);
                                    }
                                    return new IntCell((int)d);
                                } catch (NumberFormatException e) {
                                    return new MissingCell(e.getMessage());
                                }
                            }
                            if (type.isCompatible(DoubleValue.class)) {
                                try {
                                    return new DoubleCell(Double.parseDouble(valueAsString));
                                } catch (NumberFormatException e) {
                                    return new MissingCell(e.getMessage());
                                }
                            }
                            if (type.isCompatible(LocalDateTimeValue.class)) {
                                try {
                                    return LocalDateTimeCellFactory.create(valueAsString);
                                } catch (DateTimeParseException e) {
                                    try {
                                        LocalDate date = LocalDate.parse(valueAsString);
                                        return LocalDateTimeCellFactory.create(LocalDateTime.from(date.atStartOfDay()));
                                    } catch (DateTimeParseException ed) {
                                        LocalTime time = LocalTime.parse(valueAsString);
                                        return LocalDateTimeCellFactory
                                            .create(LocalDateTime.from(time.atDate(LocalDate.of(1900, 1, 1))));
                                    }
                                }
                            }
                            if (type.isCompatible(StringValue.class)) {
                                //TODO the dates are in standardized format, which does not work as probably expected.
                                return new StringCell(valueAsString);
                            }
                        } catch (RuntimeException e) {
                            return new MissingCell(e.getMessage());
                        }
                        //TODO anything missing? Should we support other types?
                        return DataType.getMissingCell();
                    }

                    private boolean isSkipRow(final Integer key) {
                        boolean isHeader = isHeader(key);
                        int lastRow = settings.getLastRow0();
                        return isHeader || (!settings.getReadAllData()
                            && (key < settings.getFirstRow0() || (lastRow >= 0 && key > lastRow)));
                    }

                    /**
                     * @param key
                     * @return
                     */
                    private boolean isHeader(final Integer key) {
                        boolean isHeader = settings.getHasColHeaders() && key == settings.getColHdrRow0();
                        return isHeader;
                    }

                    private Entry<Integer, Map<Integer, Content>> nextEntry() {
                        if (m_rowNumber <= numOfRows()) {
                            return new AbstractMap.SimpleImmutableEntry<>(m_rowNumber, constructRow(m_rowNumber++));
                        }
                        return null;
                    }

                    @Override
                    public DataRow next() {
                        if (!hasNext()) {
                            throw new NoSuchElementException();
                        }
                        m_calledNext = true;
                        return new DefaultRow(m_rowKey, m_cells);
                    }
                };
            }
        };
        if (resultExcelToKNIME != null) {
            try {
                resultExcelToKNIME.putAll(mapFromExcelColumnIndicesToKNIME.entrySet().stream()
                    .collect(Collectors.toMap(e -> e.getValue(), e -> e.getKey())));
            } catch (RuntimeException e) {
                //Probably read-only, ignore.
                e.printStackTrace();
            }
        }
        return dataTable;
    }

    /**
     * Generates a {@link DataTable} based on {@code rawSettings}.
     *
     * @param rawSettings The settings to be used to generate the table.
     * @param resultExcelToKNIME The mapping of excel {@code 1-based} columns to KNIME output columns {@code 0}-based.
     *            This will be a result.
     * @return The {@link DataTable}.
     * @throws InvalidSettingsException Some settings in {@code rawSettings} is incorrect.
     */
    DataTable createDataTable(final Path path, final XLSUserSettings rawSettings,
        final Map<Integer, Integer> resultExcelToKNIME) throws InvalidSettingsException {
        return createDataTable(path, rawSettings, resultExcelToKNIME, -1, 0, new ValueUniquifier());
    }

    /**
     * @param spec Original spec with all information present.
     * @return The updated spec with no data and time columns.
     */
    private static DataTableSpec updatedSpec(final DataTableSpec spec) {
        final DataTableSpecCreator tableSpecCreator = new DataTableSpecCreator(spec);
        for (int i = 0; i < spec.getNumColumns(); i++) {
            if (spec.getColumnSpec(i).getType().isCompatible(LocalDateTimeValue.class)) {
                tableSpecCreator.replaceColumn(i,
                    new DataColumnSpecCreator(spec.getColumnSpec(i).getName(), LocalDateTimeCellFactory.TYPE).createSpec());
            }
        }
        return tableSpecCreator.createSpec();
    }

    /**
     * Looks at the specified XLS file and tries to determine the type of all columns contained. Also returns the number
     * of columns and rows in all sheets.
     *
     * @param the path or url
     * @param settings the pre-set settings.
     * @return the result settings
     */
    private List<DataType> analyzeColumnTypes(final Path path, final XLSUserSettings settings,
        final Set<Integer> skippedCols) {

        if (settings == null) {
            throw new NullPointerException("Settings can't be null");
        }
        if (path == null) {
            throw new NullPointerException("File location must be set.");
        }

        return setColumnTypes(settings, skippedCols);
    }

    /**
     * Traverses the specified sheet in the file and detects the type for all columns in the sheets.
     *
     * @param settings The user settings
     * @param skippedCols The skipped columns (non-{@code null}, {@code 0}=based. this method fills it)
     */
    private List<DataType> setColumnTypes(final XLSUserSettings settings, final Set<Integer> skippedCols) {

        skippedCols.clear();

        String sheetName = settings.getSheetName();
        if (sheetName == null || sheetName.isEmpty() || sheetName.equals(XLSReaderNodeDialog.FIRST_SHEET)) {
            throw new IllegalStateException("At this point sheetName has to be an existing sheet name: " + sheetName);
        }
        //CombinationStrategies.KeepEverythingAsIsCombineOnlyIdentical
        return collectDataTypes(settings, skippedCols);
    }

    /**
     * @param settings The {@link XLSUserSettings} to be used.
     * @param skippedCols The skipped columns ({@code 0-based, this method fills it}.
     * @return The data types of the columns.
     */
    private List<DataType> collectDataTypes(final XLSUserSettings settings, final Set<Integer> skippedCols) {
        final List<DataType> ret = new ArrayList<>();
        int firstColumn = settings.getFirstColumn0(), lastColumn = settings.getLastColumn0();
        SortedMap<Integer, SortedMap<Integer, Pair<ActualDataType, Integer>>> dataTypes = dataTypes();
        //        System.out.println(dataTypes);
        int actualLastColumn = lastColumnIndex();
        if (actualLastColumn < 0) {
            return Collections.emptyList();
        }
        if (lastColumn < 0 && !dataTypes.isEmpty()) {
            lastColumn = dataTypes.lastKey();
        }
        for (int i = 0; i <= lastColumn; ++i) {
            if ((settings.getSkipHiddenColumns() && m_hiddenColumns.contains(i + 1))
                || (firstColumn >= 0 && i < firstColumn)
                || (settings.getHasRowHeaders() && i == settings.getRowHdrCol0())) {
                skippedCols.add(i);
            }
        }
        for (int i = 0; i <= lastColumn; ++i) {
            if (settings.getSkipHiddenColumns() && m_hiddenColumns.contains(i + 1)) {
                skippedCols.add(i);
            } else if (firstColumn >= 0 && i < firstColumn) {
                skippedCols.add(i);
            } else if (settings.getHasRowHeaders() && i == settings.getRowHdrCol0()) {
                skippedCols.add(i);
            } else if (dataTypes.get(i) == null || POIUtils.isAllEmpty(dataTypes.get(i), settings)) {
                if (settings.getSkipEmptyColumns()) {
                    skippedCols.add(i);
                } else {
                    ret.add(StringCell.TYPE);
                }
            } else {
                ret.add(dataTypeOf(dataTypes.get(i), settings));
            }
        }
        for (int i = actualLastColumn + 1; i <= lastColumn; ++i) {
            if (settings.getSkipEmptyColumns()) {
                skippedCols.add(i);
            } else {
                ret.add(StringCell.TYPE);
            }
        }
        return ret;
    }

    /**
     * Computes the column's compressed type information.
     *
     * @return A {@link SortedMap} with {@code 0}-based keys and values (both inclusive, an interval) and the type of
     *         the cells.
     */
    private SortedMap<Integer, SortedMap<Integer, Pair<ActualDataType, Integer>>> dataTypes() {
        SortedMap<Integer, SortedMap<Integer, Pair<ActualDataType, Integer>>> ret = new TreeMap<>();
        for (int i = 0; i < m_contents.length; i++) {
            ArrayList<Content> arrayList = m_contents[i];
            if (arrayList != null) {
                Content firstContent = arrayList.size() > 0 ? arrayList.get(0) : null;
                ColumnTypeCombinator combiner = new ColumnTypeCombinator(0,
                    Pair.create(firstContent == null ? ActualDataType.MISSING : firstContent.m_type, 0));
                for (int j = 1; j < arrayList.size(); j++) {
                    Content content = arrayList.get(j);
                    combiner.combine(j,
                        content == null ? ActualDataType.MISSING : content.m_type);
                }
                ret.put(i, combiner.getCurrentTypes());
            }
        }
        return ret;
    }

    /**
     * @return The last column index within the cache.
     */
    private int lastColumnIndex() {
        return m_lastColumnIndex > Integer.MIN_VALUE ? m_lastColumnIndex
            : (m_lastColumnIndex = lastNonNull(m_contents));
    }

    /**
     * @return The last row's index in case it was fully read, otherwise the {@code -(number of read rows) - 1}.
     */
    int lastRow() {
        int length = numOfRows();
        return m_incomplete ? -length - 1 : length;
    }

    /**
     * @return
     */
    private int numOfRows() {
        return m_lastRowIndex > Integer.MIN_VALUE ? m_lastRowIndex : (m_lastRowIndex = lastRow(m_contents));
    }

    /**
     * @param contents
     * @return
     */
    private static int lastRow(final ArrayList<Content>[] contents) {
        int max = 0;
        for (ArrayList<Content> arrayList : contents) {
            max = Math.max(realSize(arrayList), max);
        }
        return max;
    }

    /**
     * @param arrayList
     * @return The last cell containing value.
     */
    private static int realSize(final ArrayList<Content> arrayList) {
        if (arrayList == null) {
            return -1;
        }
        for (int size = arrayList.size(); size-- > 0;) {
            if (arrayList.get(size) != null) {
                return size;
            }
        }
        return -1;
    }

    /**
     * @param columnMap A column's type map.
     * @param settings The user settings.
     * @return The KNIME {@link DataType} of the column.
     */
    private static DataType dataTypeOf(final SortedMap<Integer, Pair<ActualDataType, Integer>> columnMap,
        final XLSUserSettings settings) {
        //Not all empty
        int firstRow = settings.getFirstRow0(), lastRow = settings.getLastRow0();

        DataType common = DataType.getMissingCell().getType();
        final int colHdrRow = settings.getColHdrRow0();
        //TODO use head and tail
        for (Entry<Integer, Pair<ActualDataType, Integer>> entry : columnMap.entrySet()) {
            if (firstRow >= 0 && entry.getValue().getSecond() < firstRow) {
                continue;
            }
            if (lastRow >= 0 && entry.getKey() > lastRow) {
                break;
            }
            if (settings.getHasColHeaders() && colHdrRow == entry.getKey()
                && colHdrRow == entry.getValue().getSecond()) {
                continue;
            }
            ActualDataType dataType = entry.getValue().getFirst();
            if (ActualDataType.isMissing(dataType)
                || (!settings.getUseErrorPattern() && ActualDataType.isError(dataType))) {
                continue;
            }
            common = combine(common, dataType, settings);
        }
        return common == DataType.getMissingCell().getType() ? /* probably should not happen */StringCell.TYPE : common;
    }

    /**
     * Combines {@link ActualDataType}s and {@link DataType}s.
     *
     * @param common The previous {@link DataType}.
     * @param dataType The new {@link ActualDataType} to combine.
     * @param settings The user settings.
     * @return The new {@link DataType}.
     */
    private static DataType combine(final DataType common, ActualDataType dataType, final XLSUserSettings settings) {
        //if (true/*!settings.isWriteFormula()*/) {
        switch (dataType) {
            case BOOLEAN_FORMULA:
                dataType = ActualDataType.BOOLEAN;
                break;
            case DATE_FORMULA:
                dataType = ActualDataType.DATE;
                break;
            case ERROR_FORMULA:
                dataType = ActualDataType.ERROR;
                break;
            case MISSING_FORMULA:
                dataType = ActualDataType.MISSING;
                break;
            case NUMBER_FORMULA:
                dataType = ActualDataType.NUMBER;
                break;
            case NUMBER_DOUBLE_FORMULA:
                dataType = ActualDataType.NUMBER_DOUBLE;
                break;
            case NUMBER_INT_FORMULA:
                dataType = ActualDataType.NUMBER_INT;
                break;
            case STRING_FORMULA:
                dataType = ActualDataType.STRING;
                break;
            case BOOLEAN:
            case DATE:
            case ERROR:
            case MISSING:
            case NUMBER:
            case NUMBER_DOUBLE:
            case NUMBER_INT:
            case STRING:
                //Nothing to change
                break;
            default:
                throw new UnsupportedOperationException("Unknown type: " + dataType);
        }
        //        }
        switch (dataType) {
            case BOOLEAN:
                if (common == BooleanCell.TYPE) {
                    return common;
                }
                if (common.isCompatible(MissingValue.class)) {
                    return BooleanCell.TYPE;
                }
                if (common == StringCell.TYPE || common == IntCell.TYPE || common == DoubleCell.TYPE
                    || common == LocalDateTimeCellFactory.TYPE) {
                    return StringCell.TYPE;
                }
                throw new IllegalStateException("common: " + common);
            case DATE:
                if (common == LocalDateTimeCellFactory.TYPE || common.isCompatible(MissingValue.class)) {
                    return LocalDateTimeCellFactory.TYPE;
                }
                if (common == StringCell.TYPE || common == IntCell.TYPE || common == DoubleCell.TYPE
                    || common == BooleanCell.TYPE) {
                    return StringCell.TYPE;
                }
                throw new IllegalStateException("common: " + common);
            case ERROR:
                if (settings.getUseErrorPattern()) {
                    return StringCell.TYPE;
                }
                return common;
            case MISSING:
                return common;
            case NUMBER:
            case NUMBER_DOUBLE:
                if (common == DoubleCell.TYPE || common == IntCell.TYPE || common.isCompatible(MissingValue.class)) {
                    return DoubleCell.TYPE;
                }
                if (common == StringCell.TYPE || common == LocalDateTimeCellFactory.TYPE || common == BooleanCell.TYPE) {
                    return StringCell.TYPE;
                }
                throw new IllegalStateException("common: " + common);
            case NUMBER_INT:
                if (common == IntCell.TYPE || common.isCompatible(MissingValue.class)) {
                    return IntCell.TYPE;
                }
                if (common == DoubleCell.TYPE) {
                    return common;
                }
                if (common == StringCell.TYPE || common == LocalDateTimeCellFactory.TYPE || common == BooleanCell.TYPE) {
                    return StringCell.TYPE;
                }
                throw new IllegalStateException("common: " + common);

            case STRING:
                return StringCell.TYPE;
            default:
                throw new IllegalStateException("common: " + common + " Formulae are not yet supported: " + dataType);
        }
    }

    /**
     * Creates and returns a new table spec from the current settings.
     *
     * @return a new table spec from the current settings
     * @throws InvalidSettingsException if settings are invalid.
     */
    private DataTableSpec createSpec(final Path path, final XLSUserSettings settings, final Set<Integer> skippedCols) {

        List<DataType> columnTypes = analyzeColumnTypes(path, settings, skippedCols);

        int numOfCols = columnTypes.size();
        String[] colHdrs = createColHeaders(settings, numOfCols, skippedCols);
        assert colHdrs.length == numOfCols;

        DataColumnSpec[] colSpecs = new DataColumnSpec[numOfCols];

        for (int col = 0; col < numOfCols; col++) {
            assert(col < columnTypes.size() && columnTypes.get(col) != null);
            colSpecs[col] = new DataColumnSpecCreator(colHdrs[col], columnTypes.get(col)).createSpec();
        }

        // create a name
        String sheetName = settings.getSheetName();
        if (sheetName == null || sheetName.isEmpty()) {
            throw new IllegalStateException("Sheet name should be valid at this point: " + sheetName);
        }

        String tableName = path.toString() + " [" + sheetName + "]";
        return new DataTableSpec(tableName, colSpecs);
    }

    /**
     * Either reads the column names from the sheet, of generates new ones.
     *
     */
    private String[] createColHeaders(final XLSUserSettings settings, final int numOfCols,
        final Set<Integer> skippedCols) {

        String[] result = null;
        if (settings.getHasColHeaders() && !settings.getKeepXLColNames()) {
            result = readColumnHeaders(settings, numOfCols, skippedCols);
        }
        if (result == null) {
            result = new String[numOfCols];
        }
        POIUtils.fillEmptyColHeaders(settings, skippedCols, result);
        return result;
    }

    /**
     * Reads the contents of the specified column header row. Names not specified in the sheet are null in the result.
     * Uniquifies the names from the sheet. The length of the returned array is determined by the first and last column
     * to read. The result is null, if the hasColHdr flag is not set, or the specified header row is not in the sheet.
     *
     * @param settings valid settings
     * @param numOfCols number of columns to read headers for
     * @param skippedCols columns in the sheet being skipped (due to emptiness or row headers)
     * @return Returns null, if sheet contains no column headers, or the specified header row is not in the sheet,
     *         otherwise an array with unique names, or null strings, if the sheet didn't contain a name for that
     *         column.
     * @throws InvalidSettingsException if settings are invalid
     */
    private String[] readColumnHeaders(final XLSUserSettings settings, final int numOfCols,
        final Set<Integer> skippedCols) {

        String[] result = new String[numOfCols];

        if (!settings.getHasColHeaders()) {
            return result;
        }
        final Map<Integer, Content> colHeader = constructRow(settings.getColHdrRow0());
        //        if (colHeader == null) {
        //            // table had too few rows
        //            LOGGER.warn("Specified column header row not contained " + "in sheet");
        //            return result;
        //        }
        //        if (numOfCols <= colHeader.getNumCells()) {
        //            LOGGER.debug("The column header row just read has " + row.getNumCells() + " cells, but we expected "
        //                + numOfCols);
        //        }
        final ValueUniquifier uniquifier = new ValueUniquifier();
        colHeader.keySet().stream().max(Comparator.naturalOrder()).ifPresent((max) -> {
            //Necessary to skip the same columns.
            int idx = 0;
            for (int i = 0; i <= max && idx < result.length; ++i) {
                if (!skippedCols.contains(i)) {
                    if (colHeader.containsKey(i)) {
                        final Content content = colHeader.get(i);
                        final String m_valueAsString = content == null ? null : content.m_valueAsString;
                        result[idx++] =
                            m_valueAsString == null ? null : uniquifier.uniquifyRowHeader(m_valueAsString.trim());
                    } else {
                        result[idx++] = null;
                    }
                }
            }
        });
        return result;
    }

    /**
     *
     *
     * @param rowIndex0
     * @return
     */
    private Map<Integer, Content> constructRow(final int rowIndex0) {
        m_rowMap.clear();
        if (rowIndex0 < 0) {
            return m_rowMap;
        }
        for (int i = 0; i <= lastColumnIndex(); i++) {
            ArrayList<Content> list = m_contents[i];
            if (list != null) {
                Content content = list.size() > rowIndex0 ? list.get(rowIndex0) : null;
                if (content != null) {
                    m_rowMap.put(i, content);
                }
            }
        }
        return m_rowMap;
    }
}
