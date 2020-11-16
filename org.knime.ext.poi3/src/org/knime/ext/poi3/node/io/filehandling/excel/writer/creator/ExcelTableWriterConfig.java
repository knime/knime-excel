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
 *   Nov 16, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.creator;

import java.util.EnumSet;

import org.apache.poi.ss.usermodel.Workbook;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.config.AbstractExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.config.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.WorkbookCreator;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

/**
 * The configuration of the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterConfig extends AbstractExcelTableConfig<SettingsModelWriterFileChooser> {

    private static final ExcelFormat DEFAULT_EXCEL_FORMAT = ExcelFormat.XLSX;

    private static final String CFG_EXCEL_FORMAT = "excel_format";

    private final SettingsModelString m_excelFormat;

    /**
     * Constructor.
     *
     * @param portsCfg the {@link PortsConfiguration}
     */
    ExcelTableWriterConfig(final PortsConfiguration portsCfg) {
        super(portsCfg,
            new SettingsModelWriterFileChooser(CFG_FILE_CHOOSER, portsCfg,
                ExcelTableWriterNodeFactory.FS_CONNECT_GRP_ID, FilterMode.FILE, FileOverwritePolicy.FAIL,
                EnumSet.of(FileOverwritePolicy.FAIL, FileOverwritePolicy.OVERWRITE),
                DEFAULT_EXCEL_FORMAT.getFileExtension()),
            ExcelTableWriterNodeFactory.SHEET_GRP_ID);
        m_excelFormat = new SettingsModelString(CFG_EXCEL_FORMAT, DEFAULT_EXCEL_FORMAT.name());
    }

    @Override
    public boolean evaluate() {
        return false;
    }

    @Override
    public boolean abortIfSheetExists() {
        return false;
    }

    WorkbookCreator getWorkbookCreator() {
        return new WriteWorkbookCreator(getExcelFormat());
    }

    SettingsModelString getExcelFormatModel() {
        return m_excelFormat;
    }

    ExcelFormat getExcelFormat() {
        return ExcelFormat.valueOf(m_excelFormat.getStringValue());
    }

    @Override
    protected void validateAdditionalSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_excelFormat.validateSettings(settings);
    }

    @Override
    protected void loadAdditionalSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_excelFormat.loadSettingsFrom(settings);
    }

    @Override
    protected void saveAdditionalSettingsForModel(final NodeSettingsWO settings) {
        m_excelFormat.saveSettingsTo(settings);
    }

    /**
     * {@link WorkbookCreator} for new excel files.
     *
     * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
     */
    private static class WriteWorkbookCreator implements WorkbookCreator {

        private final ExcelFormat m_format;

        /**
         * Constructor.
         *
         * @param format the {@link ExcelFormat} of the workbook to be created
         */
        public WriteWorkbookCreator(final ExcelFormat format) {
            m_format = format;
        }

        @Override
        public Workbook createWorkbook() {
            return m_format.getWorkbook();
        }

        @Override
        public ExcelTableWriter createTableWriter(final ExcelTableConfig cfg,
            final ExcelCellWriterFactory cellWriterFactory) {
            return m_format.createWriter(cfg, cellWriterFactory);
        }

    }
}
