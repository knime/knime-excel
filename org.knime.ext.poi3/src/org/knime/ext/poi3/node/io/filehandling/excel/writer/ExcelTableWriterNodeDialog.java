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
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

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
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.SharedIcons;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetNameExistsHandling;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.DialogComponentWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.util.GBCBuilder;
import org.xml.sax.SAXException;

/**
 * The dialog of the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterNodeDialog extends NodeDialogPane {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelTableWriterNodeDialog.class);

    private static final String SCANNING = "/* scanning... */";

    private final AtomicLong m_updateSheetListId = new AtomicLong(0);

    private SheetUpdater m_updateSheet = null;

    private static final Pattern FILE_EXTENSION_PATTERN = Pattern.compile(//
        Arrays.stream(ExcelFormat.values())//
            .map(ExcelFormat::getFileExtension)//
            .collect(Collectors.joining("|", "(", ")\\s*$")),
        Pattern.CASE_INSENSITIVE);

    private final ExcelTableWriterConfig m_cfg;

    private final DialogComponentStringSelection m_excelType;

    private final DialogComponentWriterFileChooser m_fileChooser;

    private final JComboBox<String>[] m_sheetNames;

    private final DialogComponentButtonGroup m_sheetNameCollisionHandling;

    private final DialogComponentBoolean m_writeRowKey;

    private final DialogComponentBoolean m_writeColHeader;

    private final DialogComponentBoolean m_skipColumnHeaderOnAppend;

    private final DialogComponentBoolean m_replaceMissings;

    private final DialogComponentString m_missingValPattern;

    private final DialogComponentBoolean m_evaluateFormulas;

    private final DialogComponentButtonGroup m_landscape;

    private final DialogComponentBoolean m_autoSize;

    private final JComboBox<PaperSize> m_paperSize;

    private final DialogComponentBoolean m_openFileAfterExec;

    private final JLabel m_openFileAfterExecLbl;

    /**
     * Constructor.
     *
     * @param portsConfig the ports configuration
     */
    @SuppressWarnings("unchecked")
    ExcelTableWriterNodeDialog(final PortsConfiguration portsConfig) {
        m_cfg = new ExcelTableWriterConfig(portsConfig);

        SettingsModelString excelFormatModel = m_cfg.getExcelFormatModel();
        m_excelType = new DialogComponentStringSelection(excelFormatModel, "Excel format    ",
            Arrays.stream(ExcelFormat.values())//
                .map(ExcelFormat::name)//
                .collect(Collectors.toList()));

        final SettingsModelWriterFileChooser writerModel = m_cfg.getFileChooserModel();
        final var writeFvm =
            createFlowVariableModel(writerModel.getKeysForFSLocation(), FSLocationVariableType.INSTANCE);
        m_fileChooser = new DialogComponentWriterFileChooser(writerModel, "excel_reader_writer", writeFvm);

        m_sheetNames = Stream.generate(JComboBox<String>::new)//
            .limit(portsConfig.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length)//
            .toArray(JComboBox[]::new);
        Arrays.stream(m_sheetNames).forEach(b -> b.setEditable(true));
        m_sheetNameCollisionHandling = new DialogComponentButtonGroup(m_cfg.getSheetExistsHandlingModel(), null, false,
            SheetNameExistsHandling.values());

        m_writeColHeader = new DialogComponentBoolean(m_cfg.getWriteColHeaderModel(), "Write column headers");
        m_skipColumnHeaderOnAppend = new DialogComponentBoolean(m_cfg.getSkipColumnHeaderOnAppendModel(),
            "Don't write column headers if sheet exists");
        m_cfg.getWriteColHeaderModel().addChangeListener(
            e -> m_cfg.getSkipColumnHeaderOnAppendModel().setEnabled(m_cfg.getWriteColHeaderModel().getBooleanValue()));

        m_writeRowKey = new DialogComponentBoolean(m_cfg.getWriteRowKeyModel(), "Write row key");
        m_replaceMissings = new DialogComponentBoolean(m_cfg.getReplaceMissingsModel(), "Replace missing values by");
        m_missingValPattern = new DialogComponentString(m_cfg.getMissingValPatternModel(), null);

        m_evaluateFormulas = new DialogComponentBoolean(m_cfg.getEvaluateFormulasModel(), "Evaluate formulas (leave unchecked if uncertain; see node description for details)");

        m_autoSize = new DialogComponentBoolean(m_cfg.getAutoSizeModel(), "Autosize columns");
        m_landscape = new DialogComponentButtonGroup(m_cfg.getLandscapeModel(), null, false, Orientation.values());
        m_paperSize = new JComboBox<>(PaperSize.values());
        m_openFileAfterExec =
            new DialogComponentBoolean(m_cfg.getOpenFileAfterExecModel(), "Open file after execution");
        m_openFileAfterExecLbl = new JLabel("");

        writerModel.addChangeListener(l -> toggleOptions());
        excelFormatModel.addChangeListener(l -> updateLocation());

        addTab("Settings", createSettings());
    }

    private void toggleOptions() {
        toggleAppendRelatedOptions();
        toggleOpenFileAfterExecOption();
        updateSheetListAndSelect();
    }

    private void toggleAppendRelatedOptions() {
        final boolean enabled = m_fileChooser.getSettingsModel().getFileOverwritePolicy() == FileOverwritePolicy.APPEND;
        m_evaluateFormulas.getModel().setEnabled(enabled);
        m_sheetNameCollisionHandling.getModel().setEnabled(enabled);
    }

    private void toggleOpenFileAfterExecOption() {
        // cannot be headless as we'd not have a dialog in this case
        final boolean isRemote = ExcelTableWriterNodeModel.isHeadlessOrRemote();
        final boolean categorySupported = ExcelTableWriterNodeModel
            .categoryIsSupported(m_fileChooser.getSettingsModel().getLocation().getFSCategory());
        m_openFileAfterExec.getModel().setEnabled(!isRemote && categorySupported);
        if (isRemote) {
            m_openFileAfterExecLbl.setIcon(SharedIcons.INFO_BALLOON.get());
            m_openFileAfterExecLbl.setText("Not support in remote job view");
        } else if (!categorySupported) {
            m_openFileAfterExecLbl.setIcon(SharedIcons.INFO_BALLOON.get());
            m_openFileAfterExecLbl.setText("Not support by the selected file system");
        } else {
            m_openFileAfterExecLbl.setIcon(null);
            m_openFileAfterExecLbl.setText("");
        }
    }

    private void updateLocation() {
        final var format = ExcelFormat.valueOf(m_cfg.getExcelFormatModel().getStringValue());
        SettingsModelWriterFileChooser writerModel = m_fileChooser.getSettingsModel();
        if (!writerModel.isOverwrittenByFlowVariable()) {
            FSLocation location = writerModel.getLocation();
            final String locPath = location.getPath();
            final String newPath = FILE_EXTENSION_PATTERN.matcher(locPath).replaceAll(format.getFileExtension());
            if (!newPath.equals(locPath)) {
                writerModel.setLocation(
                    new FSLocation(location.getFSCategory(), location.getFileSystemSpecifier().orElse(null), newPath));
            }
        }
        writerModel.setFileExtensions(format.getFileExtension());
    }

    private Component createSettings() {
        final var p = new JPanel(new GridBagLayout());
        final var gbc = new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(1).fillHorizontal().setWidth(2);
        p.add(createFileChooserPanel(), gbc.build());

        gbc.incY();
        p.add(createSheetNamesPanel(), gbc.build());

        gbc.incY();
        p.add(createNameIdPanel(), gbc.build());

        gbc.incY();
        p.add(createMissingsPanel(), gbc.build());

        gbc.incY();
        p.add(createFormulasPanel(), gbc.build());

        gbc.incY();
        p.add(createSizePanel(), gbc.build());

        gbc.incY().setWeightX(0).fillNone().setWidth(1);
        p.add(m_openFileAfterExec.getComponentPanel(), gbc.build());

        gbc.incX();
        p.add(m_openFileAfterExecLbl, gbc.build());

        gbc.resetX().incY().setWidth(2).weight(1, 1).fillBoth();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createFileChooserPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "File format & output location"));
        final var gbc = new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        p.add(m_excelType.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal().insetLeft(4);
        p.add(m_fileChooser.getComponentPanel(), gbc.build());

        return p;
    }

    private Component createSheetNamesPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Sheets"));
        final var gbc = new GBCBuilder().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        for (var i = 0; i < m_sheetNames.length; i++) {
            gbc.resetX().incY().insetLeft(4);
            p.add(new JLabel((i + 1) + ". sheet name"), gbc.build());

            gbc.incX().insetLeft(15);
            p.add(m_sheetNames[i], gbc.build());
            gbc.insetTop(3);

        }

        gbc.incY().resetX().insetLeft(4);
        p.add(new JLabel("If sheet exists"), gbc.build());

        gbc.incX().insetLeft(5);
        p.add(m_sheetNameCollisionHandling.getComponentPanel(), gbc.build());

        gbc.resetX().setWeightX(1).setWidth(2).incY().insetLeft(0).insetTop(0).fillHorizontal();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createNameIdPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Names and IDs"));
        final var gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_writeRowKey.getComponentPanel(), gbc.build());

        gbc.incY().insetTop(-5);
        p.add(m_writeColHeader.getComponentPanel(), gbc.build());
        gbc.incY();
        p.add(m_skipColumnHeaderOnAppend.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createMissingsPanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Missing value handling"));

        final var gbc =
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

        final var gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_evaluateFormulas.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal().insetTop(-10);
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createSizePanel() {
        final var p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Layout"));

        final var gbc = new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone()
            .setWidth(2).insetLeft(-3);
        p.add(m_autoSize.getComponentPanel(), gbc.build());

        gbc.incY().insetTop(-5).setWidth(1);
        p.add(m_landscape.getComponentPanel(), gbc.build());

        gbc.incX().insetLeft(12);
        p.add(m_paperSize, gbc.build());

        gbc.incY().setWidth(2).setWeightX(1).fillHorizontal();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
        m_excelType.saveSettingsTo(settings);
        m_fileChooser.saveSettingsTo(settings);
        ExcelTableWriterConfig.saveSheetNames(settings, Arrays.stream(m_sheetNames)//
            .map(JComboBox::getSelectedItem)//
            .map(String.class::cast)//
            .map(String::trim)//
            .toArray(String[]::new));
        m_sheetNameCollisionHandling.saveSettingsTo(settings);
        m_writeRowKey.saveSettingsTo(settings);
        m_writeColHeader.saveSettingsTo(settings);
        m_skipColumnHeaderOnAppend.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
        m_autoSize.saveSettingsTo(settings);
        m_landscape.saveSettingsTo(settings);
        final SettingsModelString paperSizeModel = m_cfg.getPaperSizeModel();
        paperSizeModel.setStringValue(((PaperSize)m_paperSize.getSelectedItem()).name());
        paperSizeModel.saveSettingsTo(settings);
        m_openFileAfterExec.saveSettingsTo(settings);
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        m_excelType.loadSettingsFrom(settings, specs);
        m_fileChooser.loadSettingsFrom(settings, specs);
        m_cfg.loadSheetsInDialog(settings);
        IntStream.range(0, m_sheetNames.length)//
            .forEach(i -> m_sheetNames[i].setSelectedItem(m_cfg.getSheetNames()[i]));
        m_sheetNameCollisionHandling.loadSettingsFrom(settings, specs);
        m_writeRowKey.loadSettingsFrom(settings, specs);
        m_writeColHeader.loadSettingsFrom(settings, specs);
        try {
            m_cfg.getSkipColumnHeaderOnAppendModel().loadSettingsFrom(settings);
            m_skipColumnHeaderOnAppend.loadSettingsFrom(settings, specs);
        } catch (InvalidSettingsException e) { // NOSONAR we ant to load the default here
            m_cfg.getSkipColumnHeaderOnAppendModel().setBooleanValue(true);
        }
        m_cfg.getSkipColumnHeaderOnAppendModel().setEnabled(m_cfg.getWriteColHeaderModel().getBooleanValue());
        m_replaceMissings.loadSettingsFrom(settings, specs);
        m_missingValPattern.loadSettingsFrom(settings, specs);
        m_evaluateFormulas.loadSettingsFrom(settings, specs);
        m_autoSize.loadSettingsFrom(settings, specs);
        m_landscape.loadSettingsFrom(settings, specs);
        final SettingsModelString paperSizeModel = m_cfg.getPaperSizeModel();
        try {
            paperSizeModel.loadSettingsFrom(settings);
        } catch (InvalidSettingsException e) { // NOSONAR we want to load a default here
            paperSizeModel.setStringValue(ExcelConstants.DEFAULT_PAPER_SIZE.name());
        }
        m_paperSize.setSelectedItem(PaperSize.valueOf(paperSizeModel.getStringValue()));
        m_openFileAfterExec.loadSettingsFrom(settings, specs);

        toggleOptions();
        updateLocation();
    }

    /**
     * Reads from the currently selected file the list of worksheets (in a background thread) and selects the provided
     * sheet (if not null - otherwise selects the first name). Calls {@link #sheetNameChanged()} after the update.
     */
    private void updateSheetListAndSelect() {
        final var sheetNames = m_cfg.getSheetNames();
        for (var i = 0; i < m_sheetNames.length; i++) {
            final var item = (String)m_sheetNames[i].getSelectedItem();
            if (!SCANNING.equals(item)) {
                sheetNames[i] = (String)m_sheetNames[i].getSelectedItem();
            }
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
            try (final var accessor = m_fileChooser.getSettingsModel().createWritePathAccessor()) {
                final var path = accessor.getOutputPath(x -> LOGGER.error(x.getMessage()));
                if (!Files.exists(path)) {
                    return Collections.emptyMap();
                }

                final var excelFormat = ExcelFormat.valueOf(m_cfg.getExcelFormatModel().getStringValue());

                switch (excelFormat) {
                    case XLS:
                        final var bufferedInputStream = new BufferedInputStream(Files.newInputStream(path));
                        try (final var workbook = WorkbookFactory.create(bufferedInputStream)) {
                            return ExcelUtils.getSheetNames(workbook);
                        }
                    case XLSX:
                        return readXLSXSheets(path);
                    default:
                        throw new InvalidSettingsException("Only .xls and .xlsx file formats are allowed");
                }
            }
        }

        private Map<String, Boolean> readXLSXSheets(final Path path)
            throws IOException, SAXException, ParserConfigurationException, OpenXML4JException {
            try (final var opc = OPCPackage.open(Files.newInputStream(path))) {
                final var xssfReader = new XSSFReader(opc);
                final var xmlReader = XMLHelper.newXMLReader();
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
                if (selectedSheet != null) {
                    m_sheetNames[i].setSelectedItem(selectedSheet);
                } else {
                    m_sheetNames[i].setSelectedIndex(!sheetNames.isEmpty() ? 0 : -1);
                }
            }
        }
    }

    @Override
    public void onClose() {
        m_fileChooser.onClose();
        if (m_updateSheet != null && !m_updateSheet.isDone()) {
            m_updateSheet.cancel(true);
        }
    }

    private static void fixInterrupt() {
        final var currentThread = Thread.currentThread();
        if (currentThread.isInterrupted()) {
            currentThread.interrupt();
        }
    }

}
