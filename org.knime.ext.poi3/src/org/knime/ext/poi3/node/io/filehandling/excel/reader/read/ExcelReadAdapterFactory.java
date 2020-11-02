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
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Collections;
import java.util.EnumMap;
import java.util.Map;

import org.knime.core.data.DataType;
import org.knime.core.data.convert.map.BooleanCellValueProducer;
import org.knime.core.data.convert.map.DoubleCellValueProducer;
import org.knime.core.data.convert.map.IntCellValueProducer;
import org.knime.core.data.convert.map.LongCellValueProducer;
import org.knime.core.data.convert.map.MappingException;
import org.knime.core.data.convert.map.MappingFramework;
import org.knime.core.data.convert.map.PrimitiveCellValueProducer;
import org.knime.core.data.convert.map.ProducerRegistry;
import org.knime.core.data.convert.map.SimpleCellValueProducerFactory;
import org.knime.core.data.def.BooleanCell;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.LongCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.data.time.localdate.LocalDateCellFactory;
import org.knime.core.data.time.localdatetime.LocalDateTimeCellFactory;
import org.knime.core.data.time.localtime.LocalTimeCellFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.node.table.reader.ReadAdapter;
import org.knime.filehandling.core.node.table.reader.ReadAdapter.ReadAdapterParams;
import org.knime.filehandling.core.node.table.reader.ReadAdapterFactory;

/**
 * Factory for {@link ExcelReadAdapter} objects.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 * @author Adrian Nembach, KNIME GmbH, Konstanz, Germany
 */
@SuppressWarnings("javadoc")
public enum ExcelReadAdapterFactory implements ReadAdapterFactory<KNIMECellType, ExcelCell> {

        /** The singleton instance. */
        INSTANCE;

    private static final ProducerRegistry<KNIMECellType, ExcelReadAdapter> PRODUCER_REGISTRY =
        initializeProducerRegistry();

    private static final Map<KNIMECellType, DataType> DEFAULT_TYPES = createDefaultTypeMap();

    private static Map<KNIMECellType, DataType> createDefaultTypeMap() {
        final Map<KNIMECellType, DataType> defaultTypes = new EnumMap<>(KNIMECellType.class);
        defaultTypes.put(KNIMECellType.BOOLEAN, BooleanCell.TYPE);
        defaultTypes.put(KNIMECellType.INT, IntCell.TYPE);
        defaultTypes.put(KNIMECellType.LONG, LongCell.TYPE);
        defaultTypes.put(KNIMECellType.DOUBLE, DoubleCell.TYPE);
        defaultTypes.put(KNIMECellType.STRING, StringCell.TYPE);
        defaultTypes.put(KNIMECellType.LOCAL_TIME, LocalTimeCellFactory.TYPE);
        defaultTypes.put(KNIMECellType.LOCAL_DATE, LocalDateCellFactory.TYPE);
        defaultTypes.put(KNIMECellType.LOCAL_DATE_TIME, LocalDateTimeCellFactory.TYPE);
        return Collections.unmodifiableMap(defaultTypes);
    }

    private static ProducerRegistry<KNIMECellType, ExcelReadAdapter> initializeProducerRegistry() {
        final ProducerRegistry<KNIMECellType, ExcelReadAdapter> registry =
            MappingFramework.forSourceType(ExcelReadAdapter.class);
        registry.register(new SupplierCellValueProducerFactory<>(KNIMECellType.INT, Integer.class,
            StringToIntCellValueProducer::new));
        registry.register(new SupplierCellValueProducerFactory<>(KNIMECellType.DOUBLE, Double.class,
            StringToDoubleCellValueProducer::new));
        registry.register(
            new SupplierCellValueProducerFactory<>(KNIMECellType.LONG, Long.class, StringToLongCellValueProducer::new));
        registry.register(new SupplierCellValueProducerFactory<>(KNIMECellType.BOOLEAN, Boolean.class,
            StringToBooleanCellValueProducer::new));
        registry.register(new SimpleCellValueProducerFactory<>(KNIMECellType.STRING, String.class,
            ExcelReadAdapterFactory::readStringFromSource));
        registry.register(new SimpleCellValueProducerFactory<>(KNIMECellType.LOCAL_DATE, LocalDate.class,
            ExcelReadAdapterFactory::readLocalDateFromSource));
        registry.register(new SimpleCellValueProducerFactory<>(KNIMECellType.LOCAL_TIME, LocalTime.class,
            ExcelReadAdapterFactory::readLocalTimeFromSource));
        registry.register(new SimpleCellValueProducerFactory<>(KNIMECellType.LOCAL_DATE_TIME, LocalDateTime.class,
            ExcelReadAdapterFactory::readLocalDateTimeFromSource));
        return registry;
    }

    private static String readStringFromSource(final ExcelReadAdapter source,
        final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) {
        final ExcelCell excelCell = source.get(params);
        return excelCell == null ? null : excelCell.getStringValue();
    }

    private static LocalDate readLocalDateFromSource(final ExcelReadAdapter source,
        final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) {
        final ExcelCell excelCell = source.get(params);
        return excelCell == null ? null : excelCell.getLocalDateValue();
    }

    private static LocalTime readLocalTimeFromSource(final ExcelReadAdapter source,
        final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) {
        final ExcelCell excelCell = source.get(params);
        return excelCell == null ? null : excelCell.getLocalTimeValue();
    }

    private static LocalDateTime readLocalDateTimeFromSource(final ExcelReadAdapter source,
        final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) {
        final ExcelCell excelCell = source.get(params);
        return excelCell == null ? null : excelCell.getLocalDateTimeValue();
    }

    private abstract static class AbstractReadAdapterToPrimitiveCellValueProducer<S extends ReadAdapter<?, ?>, T>
        implements PrimitiveCellValueProducer<S, T, ReadAdapterParams<S, ExcelTableReaderConfig>> {

        @Override
        public final boolean producesMissingCellValue(final S source,
            final ReadAdapterParams<S, ExcelTableReaderConfig> params) throws MappingException {
            return source.get(params) == null;
        }
    }

    private static class StringToIntCellValueProducer
        extends AbstractReadAdapterToPrimitiveCellValueProducer<ExcelReadAdapter, Integer>
        implements IntCellValueProducer<ExcelReadAdapter, ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig>> {

        @Override
        public int produceIntCellValue(final ExcelReadAdapter source,
            final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) throws MappingException {
            return source.get(params).getIntValue();
        }

    }

    private static class StringToDoubleCellValueProducer
        extends AbstractReadAdapterToPrimitiveCellValueProducer<ExcelReadAdapter, Double> implements
        DoubleCellValueProducer<ExcelReadAdapter, ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig>> {

        @Override
        public double produceDoubleCellValue(final ExcelReadAdapter source,
            final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) throws MappingException {
            return source.get(params).getDoubleValue();
        }

    }

    private static class StringToLongCellValueProducer
        extends AbstractReadAdapterToPrimitiveCellValueProducer<ExcelReadAdapter, Long> implements
        LongCellValueProducer<ExcelReadAdapter, ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig>> {

        @Override
        public long produceLongCellValue(final ExcelReadAdapter source,
            final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) throws MappingException {
            return source.get(params).getLongValue();
        }

    }

    private static class StringToBooleanCellValueProducer
        extends AbstractReadAdapterToPrimitiveCellValueProducer<ExcelReadAdapter, Boolean> implements
        BooleanCellValueProducer<ExcelReadAdapter, ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig>> {

        @Override
        public boolean produceBooleanCellValue(final ExcelReadAdapter source,
            final ReadAdapterParams<ExcelReadAdapter, ExcelTableReaderConfig> params) throws MappingException {
            return source.get(params).getBooleanValue();
        }

    }

    @Override
    public ReadAdapter<KNIMECellType, ExcelCell> createReadAdapter() {
        return new ExcelReadAdapter();
    }

    @Override
    public ProducerRegistry<KNIMECellType, ExcelReadAdapter> getProducerRegistry() {
        return PRODUCER_REGISTRY;
    }

    @Override
    public Map<KNIMECellType, DataType> getDefaultTypeMap() {
        return DEFAULT_TYPES;
    }

}
