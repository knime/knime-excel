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
package org.knime.ext.poi3.node.io.filehandling.excel.writer.appender;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.util.CheckUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.cell.ExcelCellWriterFactory;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.config.AbstractExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.config.ExcelTableConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.ExcelTableWriter;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.table.WorkbookCreator;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

/**
 * The configuration of the 'Excel Table Appender' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableAppenderConfig extends AbstractExcelTableConfig<SettingsModelReaderFileChooser> {

    private static final String[] FILE_EXTENSION =
        Arrays.stream(ExcelFormat.values()).map(ExcelFormat::getFileExtension).toArray(String[]::new);

    private static final String CFG_FAIL_IF_SHEET_EXISTS = "abort_if_already_sheet_exists";

    private static final String CFG_EVALUATE_FORMULAS = "evaluate_formulas";

    private final SettingsModelBoolean m_abortIfSheetExists;

    private final SettingsModelBoolean m_evaluateFormulas;

    /**
     * Constructor.
     *
     * @param portsCfg the {@link PortsConfiguration}
     */
    ExcelTableAppenderConfig(final PortsConfiguration portsCfg) {
        super(portsCfg,
            new SettingsModelReaderFileChooser(CFG_FILE_CHOOSER, portsCfg,
                ExcelTableAppenderNodeFactory.FS_CONNECT_GRP_ID, FilterMode.FILE, FILE_EXTENSION),
            ExcelTableAppenderNodeFactory.SHEET_GRP_ID);
        m_abortIfSheetExists = new SettingsModelBoolean(CFG_FAIL_IF_SHEET_EXISTS, true);
        m_evaluateFormulas = new SettingsModelBoolean(CFG_EVALUATE_FORMULAS, false);
    }

    SettingsModelBoolean getAbortIfSheetExistsModel() {
        return m_abortIfSheetExists;
    }

    SettingsModelBoolean getEvaluateFormulasModel() {
        return m_evaluateFormulas;
    }

    @Override
    public boolean evaluate() {
        return m_evaluateFormulas.getBooleanValue();
    }

    @Override
    public boolean abortIfSheetExists() {
        return m_abortIfSheetExists.getBooleanValue();
    }

    static WorkbookCreator getWorkbookCreator(final FSPath input) {
        return new AppendWorkbookCreator(input);
    }

    @Override
    protected void validateAdditionalSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_abortIfSheetExists.validateSettings(settings);
        m_evaluateFormulas.validateSettings(settings);
    }

    @Override
    protected void loadAdditionalSettingsForModel(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_abortIfSheetExists.loadSettingsFrom(settings);
        m_evaluateFormulas.loadSettingsFrom(settings);
    }

    @Override
    protected void saveAdditionalSettingsForModel(final NodeSettingsWO settings) {
        m_abortIfSheetExists.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
    }

    /**
     * {@link WorkbookCreator} for existing excel files.
     *
     * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
     */
    private static class AppendWorkbookCreator implements WorkbookCreator {

        private final FSPath m_path;

        private ExcelFormat m_format;

        /**
         * Constructor.
         *
         * @param path path to the excel file
         */
        AppendWorkbookCreator(final FSPath path) {
            m_path = path;
        }

        @SuppressWarnings("resource")
        @Override
        public Workbook createWorkbook() throws IOException {
            // if create fails the input stream gets closed
            final Workbook wb = WorkbookFactory.create(new BufferedInputStream(Files.newInputStream(m_path)));
            if (wb instanceof HSSFWorkbook) {
                m_format = ExcelFormat.XLS;
            } else if (wb instanceof XSSFWorkbook) {
                m_format = ExcelFormat.XLSX;
            } else {
                wb.close();
                throw new IllegalArgumentException(
                    String.format("Unsupported format: Unable to append spreadsheets to %s", m_path.toString()));
            }
            return wb;
        }

        @Override
        public ExcelTableWriter createTableWriter(final ExcelTableConfig cfg,
            final ExcelCellWriterFactory cellWriterFactory) {
            CheckUtils.checkState(m_format != null, "Cannot create a table writer before creating a workbook");
            return m_format.createWriter(cfg, cellWriterFactory);
        }

    }
}
