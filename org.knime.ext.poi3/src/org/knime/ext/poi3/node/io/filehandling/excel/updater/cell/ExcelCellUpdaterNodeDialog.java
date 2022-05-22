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
 *   Nov 6, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.updater.cell;

import java.awt.Component;
import java.awt.GridBagLayout;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.StringValue;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.DialogComponent;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentColumnNameSelection;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.DialogComponentReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.DialogComponentWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.util.GBCBuilder;
import org.xml.sax.SAXException;

/**
 * The dialog of the 'Excel Cell Updater' node.
 *
 * @author Moditha Hewasinghage, KNIME GmbH, Berlin, Germany
 * @author Jannik Löscher, KNIME GmbH, Konstanz, Germany
 */
final class ExcelCellUpdaterNodeDialog extends NodeDialogPane {

    private static final String SCANNING = "/* scanning... */";

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelCellUpdaterNodeDialog.class);

    private final int[] m_dataPortIndices;

    private final ExcelCellUpdaterConfig m_cfg;

    private final DialogComponentReaderFileChooser m_srcFileChooser;

    private final DialogComponentBoolean m_createNewFile;

    private final DialogComponentWriterFileChooser m_destFileChooser;

    private final JComboBox<String>[] m_sheetNames;

    private final DialogComponentColumnNameSelection[] m_coordinateColumns;

    private final AtomicLong m_updateSheetListId = new AtomicLong(0);

    private SheetUpdater m_updateSheet = null;

    private final JLabel m_previewMsg = new JLabel();

    private final DialogComponentBoolean m_evaluateFormulas;

    private final DialogComponentBoolean m_replaceMissings;

    private final DialogComponentString m_missingValPattern;

    /**
     * Constructor.
     *
     * @param portsConfig the ports configuration
     */
    @SuppressWarnings("unchecked")
    ExcelCellUpdaterNodeDialog(final PortsConfiguration portsConfig) {
        m_cfg = new ExcelCellUpdaterConfig(portsConfig);

        final SettingsModelReaderFileChooser readerModel = m_cfg.getSrcFileChooserModel();
        final var readFvm =
            createFlowVariableModel(readerModel.getKeysForFSLocation(), FSLocationVariableType.INSTANCE);
        m_srcFileChooser = new DialogComponentReaderFileChooser(readerModel, "excel_reader_writer", readFvm);

        m_createNewFile = new DialogComponentBoolean(m_cfg.isCreateNewFile(), "Create new file");

        final SettingsModelWriterFileChooser writerModel = m_cfg.getDestFileChooserModel();
        final var writeFvm =
            createFlowVariableModel(writerModel.getKeysForFSLocation(), FSLocationVariableType.INSTANCE);
        m_destFileChooser = new DialogComponentWriterFileChooser(writerModel, "excel_reader_writer", writeFvm);

        m_dataPortIndices = portsConfig.getInputPortLocation().get(ExcelCellUpdaterNodeFactory.SHEET_GRP_ID);

        m_coordinateColumns = Arrays.stream(m_dataPortIndices)//
            .mapToObj(i -> new DialogComponentColumnNameSelection(new SettingsModelString("x", ""),
                "Based on address column", i, StringValue.class))//
            .toArray(DialogComponentColumnNameSelection[]::new);
        m_sheetNames = Stream.generate(JComboBox<String>::new)//
            .limit(m_dataPortIndices.length)// number of input tables
            .toArray(JComboBox[]::new);

        m_evaluateFormulas = new DialogComponentBoolean(m_cfg.getEvaluateFormulasModel(), "Evaluate formulas");
        m_replaceMissings = new DialogComponentBoolean(m_cfg.getReplaceMissingsModel(), "Replace missing values by");
        m_missingValPattern = new DialogComponentString(m_cfg.getMissingValPatternModel(), null);

        addTab("Settings", createSettings());
        registerDialogChangeListeners();
        toggleDestFileChooser();
    }

    private void registerDialogChangeListeners() {
        m_srcFileChooser.getSettingsModel().addChangeListener(l -> updateSheetListAndSelect());
        m_createNewFile.getModel().addChangeListener(l -> toggleDestFileChooser());
    }

    private void toggleDestFileChooser() {
        m_destFileChooser.getSettingsModel().setEnabled(m_createNewFile.isSelected());
    }

    private Component createSettings() {
        final var p = new JPanel(new GridBagLayout());
        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(1).fillHorizontal().setWidth(2);
        p.add(createInputFilePanel(), gbc.build());

        gbc.incY();
        p.add(createOutputFilePanel(), gbc.build());

        gbc.incY();
        p.add(createUpdatePanel(), gbc.build());

        gbc.incY();
        p.add(createMissingsPanel(), gbc.build());

        gbc.incY();
        p.add(createFormulasPanel(), gbc.build());

        gbc.resetX().incY().setWidth(2).weight(1, 1).fillBoth();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createInputFilePanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Input file"));
        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        gbc.setWeightX(1).fillHorizontal().insetLeft(4);
        p.add(m_srcFileChooser.getComponentPanel(), gbc.build());
        return p;
    }

    private Component createOutputFilePanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Output file"));
        final GBCBuilder gbc = new GBCBuilder().resetX().resetY().anchorLineStart()//
            .setWeightX(0).setWeightY(0).fillNone();

        p.add(m_createNewFile.getComponentPanel(), gbc.build());
        p.add(new JPanel(), gbc.incX().fillHorizontal().setWeightX(1.0).build());
        p.add(m_destFileChooser.getComponentPanel(), gbc.resetX().setWidth(2).insetLeft(4).fillNone().incY().build());
        return p;
    }

    private Component createUpdatePanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Update"));
        final var gbc = new GBCBuilder().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        for (var i = 0; i < m_sheetNames.length; i++) {
            gbc.resetX().incY().insetLeft(4);
            p.add(new JLabel((i + 1) + ". Excel sheet"), gbc.build());

            gbc.incX().insetLeft(15);

            p.add(m_sheetNames[i], gbc.build());

            gbc.incX().insetLeft(25);
            p.add(m_coordinateColumns[i].getComponentPanel(), gbc.build());
            gbc.insetTop(3);
        }

        gbc.resetX().setWeightX(1).setWidth(3).incY().insetLeft(0).insetTop(0).fillHorizontal();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createMissingsPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Missing value handling"));

        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_replaceMissings.getComponentPanel(), gbc.build());

        gbc.incX().insetLeft(-10);
        p.add(m_missingValPattern.getComponentPanel(), gbc.build());

        gbc.incY().setWidth(2).setWeightX(1).fillHorizontal().insetTop(-10);
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createFormulasPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Formulas"));

        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_evaluateFormulas.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal().insetTop(-10);
        p.add(new JPanel(), gbc.build());

        return p;
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        m_srcFileChooser.saveSettingsTo(settings);
        m_destFileChooser.saveSettingsTo(settings);
        m_createNewFile.saveSettingsTo(settings);
        m_destFileChooser.getSettingsModel().getLocation();
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
        ExcelCellUpdaterConfig.saveSheetNames(settings, Arrays.stream(m_sheetNames)//
            .map(JComboBox::getSelectedItem)//
            .toArray(String[]::new));
        ExcelCellUpdaterConfig.saveCoordinateColumnNames(settings, Arrays.stream(m_coordinateColumns)//
            .map(DialogComponent::getModel)//
            .map(SettingsModelString.class::cast)//
            .map(SettingsModelString::getStringValue)//
            .toArray(String[]::new));
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        m_cfg.loadSheetsInDialog(settings);
        m_cfg.loadCoordinateColumnNamesInDialog(settings);
        final var names = m_cfg.getCoordinateColumnNames();
        final var dummy = new NodeSettings("dummy");
        for (var idx = 0; idx < names.length; idx++) {
            ExcelCellUpdaterNodeModel.checkExistsOrSaveFirst((DataTableSpec)specs[m_dataPortIndices[idx]], names, idx);

            dummy.addString("x", names[idx]);
            m_coordinateColumns[idx].loadSettingsFrom(dummy, specs);
        }
        m_srcFileChooser.loadSettingsFrom(settings, specs);
        m_destFileChooser.loadSettingsFrom(settings, specs);
        m_createNewFile.loadSettingsFrom(settings, specs);
        m_replaceMissings.loadSettingsFrom(settings, specs);
        m_missingValPattern.loadSettingsFrom(settings, specs);
        m_evaluateFormulas.loadSettingsFrom(settings, specs);
        updateSheetListAndSelect();
        toggleDestFileChooser();
    }

    @Override
    public void onClose() {
        m_srcFileChooser.onClose();
        m_destFileChooser.onClose();
        if (m_updateSheet != null && !m_updateSheet.isDone()) {
            m_updateSheet.cancel(true);
        }
    }

    /**
     * Reads from the currently selected file the list of worksheets (in a background thread) and selects the provided
     * sheet (if not null - otherwise selects the first name). Calls {@link #sheetNameChanged()} after the update.
     */
    private void updateSheetListAndSelect() {
        m_previewMsg.setText("Loading input file...");

        for (var i = 0; i < m_sheetNames.length; i++) {
            m_sheetNames[i].setModel(new DefaultComboBoxModel<>(new String[]{SCANNING}));
        }

        // The id of the current update
        // Note that this code and the doneWithContext is always executed by the same thread
        // Therefore we only have to make sure that the doneWithContext belongs to the most current update
        final long currentId = m_updateSheetListId.incrementAndGet();
        if (m_updateSheet != null && !m_updateSheet.isDone()) {
            m_updateSheet.cancel(true);
        }
        m_updateSheet = new SheetUpdater(currentId);
        m_updateSheet.execute();
    }

    private class SheetUpdater extends SwingWorkerWithContext<Map<String, Boolean>, Object> {

        final long m_ownId;

        /**
         * @param id the id of the updater which is used to decide whether to use the result.
         */
        SheetUpdater(final long id) {
            m_ownId = id;
        }

        @Override
        protected Map<String, Boolean> doInBackgroundWithContext() throws Exception {
            try (final var accessor = m_srcFileChooser.getSettingsModel().createReadPathAccessor()) {
                final var path = accessor.getFSPaths(x -> m_previewMsg.setText(x.getMessage())).stream().findFirst();
                if (path.isEmpty() || !Files.exists(path.get())) {
                    return Collections.emptyMap();
                }

                final var excelFormat = ExcelCellUpdaterNodeModel.getExcelFormat(path.get().getFileName().toString());

                switch (excelFormat) {
                    case XLS:
                        final var bufferedInputStream = new BufferedInputStream(Files.newInputStream(path.get()));
                        try (final var workbook = WorkbookFactory.create(bufferedInputStream)) {
                            return ExcelUtils.getSheetNames(workbook);
                        }
                    case XLSX:
                        return readXLSXSheets(path.get());
                    default:
                        throw new InvalidSettingsException("Only .xls and .xlsx file formats are allowed");
                }
            }
        }

        private Map<String, Boolean> readXLSXSheets(final Path path)
            throws IOException, SAXException, ParserConfigurationException, OpenXML4JException {
            try (final var opc = OPCPackage.open(Files.newInputStream(path))) {
                final var xssfReader = new XSSFReader(opc);
                final var xmlReader = SAXHelper.newXMLReader();
                // disable DTD to prevent almost all XXE attacks, XMLHelper.newXMLReader() did set further security features
                xmlReader.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                final var sharedStringsTable = new ReadOnlySharedStringsTable(opc, false);

                return ExcelUtils.getSheetNames(xmlReader, xssfReader, sharedStringsTable);
            }
        }

        @Override
        protected void doneWithContext() {
            if (m_ownId != m_updateSheetListId.get()) {
                // Another update of the sheet list has started
                // Do not update the sheet list
                return;
            }
            Map<String, Boolean> sheetNames = new LinkedHashMap<>();
            try {
                sheetNames = get();
            } catch (InterruptedException e) {
                fixInterrupt();
                // ignore
            } catch (CancellationException e) { // NOSONAR: only indicates closing the dialog/newer updater
                // ignore
            } catch (ExecutionException e) { // NOSONAR: We are only interested in the cause
                LOGGER.error("Could not get sheet names!", e.getCause());
            }

            for (var i = 0; i < m_sheetNames.length; i++) {
                m_sheetNames[i].setModel(new DefaultComboBoxModel<>(sheetNames.keySet().toArray(new String[0])));
                final var selectedSheet = m_cfg.getSheetNames()[i];
                if (!sheetNames.isEmpty()) {
                    if (selectedSheet != null) {
                        m_sheetNames[i].setSelectedItem(selectedSheet);
                    } else {
                        m_sheetNames[i].setSelectedIndex(0);
                    }
                } else {
                    m_sheetNames[i].setSelectedIndex(-1);
                }
            }
            updateWriterFileType();
        }

        private void updateWriterFileType() {
            final var path = Path.of(m_srcFileChooser.getSettingsModel().getLocation().getPath());
            final var excelFormat = ExcelCellUpdaterNodeModel.getExcelFormat(path.getFileName().toString());
            if (excelFormat != null) {
                SettingsModelWriterFileChooser writerModel = m_destFileChooser.getSettingsModel();
                writerModel.setFileExtensions(excelFormat.getFileExtension());
            }
        }
    }

    private static void fixInterrupt() {
        final var currentThread = Thread.currentThread();
        if (currentThread.isInterrupted()) {
            currentThread.interrupt();
        }
    }
}
