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

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.Map;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.ODFNotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.xssf.XLSBUnsupportedException;
import org.knime.core.node.ExecutionMonitor;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.WrapperExtractColumnHeaderRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xls.XLSRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.xlsx.XLSXRead;
import org.knime.filehandling.core.node.table.reader.TableReader;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.read.Read;
import org.knime.filehandling.core.node.table.reader.read.SkipIdxRead;
import org.knime.filehandling.core.node.table.reader.spec.DefaultExtractColumnHeaderRead;
import org.knime.filehandling.core.node.table.reader.spec.ExtractColumnHeaderRead;
import org.knime.filehandling.core.node.table.reader.spec.TableSpecGuesser;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderTableSpec;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TreeTypeHierarchy;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeFocusableTypeHierarchy;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeTester;

/**
 * Reader for Excel files.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class ExcelTableReader implements TableReader<ExcelTableReaderConfig, KNIMECellType, ExcelCell> {

    static final TypeFocusableTypeHierarchy<KNIMECellType, ExcelCell> TYPE_HIERARCHY = createHierarchy();

    /** The change listener that is set by the dialog to get notified once sheet names are retrieved. */
    private ChangeListener m_listener;

    /** Contains the names of the sheets as keys and whether it is the first non-empty sheet as value. */
    private Map<String, Boolean> m_sheetNames;

    @SuppressWarnings("resource") // decorated read will be closed in AbstractReadDecorator#close
    @Override
    public Read<ExcelCell> read(final Path path, final TableReadConfig<ExcelTableReaderConfig> config)
        throws IOException {
        return decorateRead(getExcelRead(path, config), config);
    }

    @SuppressWarnings("resource") // decorated read will be closed in AbstractReadDecorator#close
    @Override
    public TypedReaderTableSpec<KNIMECellType> readSpec(final Path path,
        final TableReadConfig<ExcelTableReaderConfig> config, final ExecutionMonitor exec) throws IOException {
        final TableSpecGuesser<KNIMECellType, ExcelCell> guesser = createGuesser();
        try (ExcelRead read = getExcelRead(path, config)) {
            // sheet names are already retrieved, notify a potential listener from the dialog
            m_sheetNames = read.getSheetNames();
            notifyChangeListener();
            return ExcelUtils.assignNamesIfMissing(
                guesser.guessSpec(decorateReadForSpecGuessing(read, config), config, exec), config,
                read.getHiddenColumns());
        }
    }

    @SuppressWarnings("resource") // decorated reads will be closed in AbstractReadDecorator#close
    private static Read<ExcelCell> decorateRead(final ExcelRead excelRead,
        final TableReadConfig<ExcelTableReaderConfig> config) {
        Read<ExcelCell> read = excelRead;
        if (config.useColumnHeaderIdx()) {
            read = new SkipIdxRead<>(read, config.getColumnHeaderIdx());
        }
        return ExcelUtils.decorateRowFilterReads(read, config);
    }

    @SuppressWarnings("resource") // decorated reads will be closed in AbstractReadDecorator#close
    private static ExtractColumnHeaderRead<ExcelCell> decorateReadForSpecGuessing(final ExcelRead excelRead,
        final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExtractColumnHeaderRead<ExcelCell> extractColHeaderRead =
            new DefaultExtractColumnHeaderRead<>(excelRead, config);
        final Read<ExcelCell> read = ExcelUtils.decorateRowFilterReads(extractColHeaderRead, config);
        return new WrapperExtractColumnHeaderRead(read, extractColHeaderRead::getColumnHeaders);
    }

    private static ExcelRead getExcelRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config)
        throws IOException {
        try {
            final boolean reevaluateFormulas = config.getReaderSpecificConfig().isReevaluateFormulas();
            final String pathLowerCase = path.toString().toLowerCase();
            if (!reevaluateFormulas && (pathLowerCase.endsWith(".xlsx") || pathLowerCase.endsWith(".xlsm"))) {
                return createXLSXRead(path, config);
            }
            return new XLSRead(path, config);
        } catch (ODFNotOfficeXmlFileException e) {
            // ODF (open office) files are xml files and, hence, not detected as invalid file format by the above check
            // however, ODF files are not supported
            throw createUnsupportedFileFormatException(e, path, "ODF");
        } catch (XLSBUnsupportedException e) {
            // TODO: remove this catch with AP-15391
            throw createUnsupportedFileFormatException(e, path, "XLSB");
        } catch (UnsupportedFileFormatException e) {
            throw createUnsupportedFileFormatException(e, path, null);
        } catch (IOException e) {
            // happens, e.g., with .table files
            if (e.getCause() instanceof InvalidFormatException) {
                throw createUnsupportedFileFormatException(e, path, null);
            }
            throw e;
        }
    }

    private static ExcelRead createXLSXRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config)
        throws IOException {
        try {
            return new XLSXRead(path, config);
        } catch (OLE2NotOfficeXmlFileException e) { // NOSONAR
            // Happens if an xls file has been specified that ends with xlsx or xlsm.
            // We do not fail but simply use the XLSParser instead.
            return new XLSRead(path, config);
        }
    }

    /**
     * Creates an {@link IllegalArgumentException} with a user-friendly message explaining that the specified file does
     * not have a supported format. The passed exception will be further passed into the created
     * {@link IllegalArgumentException}.
     *
     * @param e the exception to set, can be {@code null}
     * @param path the file path
     * @param fileFormat the unsupported file format if known, can be {@code null}
     * @return the created {@link IllegalArgumentException}
     */
    private static IllegalArgumentException createUnsupportedFileFormatException(final Exception e, final Path path,
        final String fileFormat) {
        final String formatString = fileFormat != null ? String.format(" (%s)", fileFormat) : "";
        return new IllegalArgumentException(
            String.format("The format%s of the file '%s' is not supported. Please select an XLSX, XLSM, or XLS file.",
                formatString, path),
            e); // TODO add XLSB with AP-15391
    }

    private static TableSpecGuesser<KNIMECellType, ExcelCell> createGuesser() {
        return new TableSpecGuesser<>(TYPE_HIERARCHY, ExcelCell::getStringValue);
    }

    private static TypeFocusableTypeHierarchy<KNIMECellType, ExcelCell> createHierarchy() {
        return TreeTypeHierarchy.builder(createTypeTester(KNIMECellType.STRING, KNIMECellType.values()))
            .addType(KNIMECellType.STRING,
                createTypeTester(KNIMECellType.DOUBLE, KNIMECellType.LONG, KNIMECellType.INT))
            .addType(KNIMECellType.DOUBLE, createTypeTester(KNIMECellType.LONG, KNIMECellType.INT))
            .addType(KNIMECellType.LONG, createTypeTester(KNIMECellType.INT))
            .addType(KNIMECellType.STRING, createTypeTester(KNIMECellType.BOOLEAN))
            .addType(KNIMECellType.STRING, createTypeTester(KNIMECellType.LOCAL_DATE_TIME, KNIMECellType.LOCAL_DATE))
            .addType(KNIMECellType.LOCAL_DATE_TIME, createTypeTester(KNIMECellType.LOCAL_DATE))
            .addType(KNIMECellType.STRING, createTypeTester(KNIMECellType.LOCAL_TIME)).build();
    }

    private static TypeTester<KNIMECellType, ExcelCell> createTypeTester(final KNIMECellType type,
        final KNIMECellType... compatibleTypes) {
        return TypeTester.createTypeTester(type,
            e -> type == e.getType() || Arrays.binarySearch(compatibleTypes, e.getType()) >= 0);
    }

    void setChangeListener(final ChangeListener l) {
        m_listener = l;
    }

    private void notifyChangeListener() {
        if (m_listener != null) {
            m_listener.stateChanged(new ChangeEvent(this));
        }
    }

    Map<String, Boolean> getSheetNames() {
        if (m_sheetNames == null) {
            return Collections.emptyMap();
        }
        return m_sheetNames;
    }

}
