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
import java.awt.Insets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.BiFunction;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;

import org.apache.poi.EncryptedDocumentException;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.DialogComponentAuthentication;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.SharedIcons;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.ext.poi3.node.io.filehandling.excel.CryptUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.DecryptionAwareWriterStatusMessageReporter;
import org.knime.ext.poi3.node.io.filehandling.excel.DialogUtil;
import org.knime.ext.poi3.node.io.filehandling.excel.StatusMessageReporterChain;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.Orientation;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.PaperSize;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.SheetNameExistsHandling;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.connections.FSLocation;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.StatusMessageReporter;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.DefaultWriterStatusMessageReporter;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.DialogComponentWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.status.DefaultStatusMessage;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage;
import org.knime.filehandling.core.util.GBCBuilder;

/**
 * The dialog of the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterNodeDialog extends NodeDialogPane {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(ExcelTableWriterNodeDialog.class);

    private static final String SCANNING = "/* scanning... */";

    private final AtomicLong m_updateSheetListId = new AtomicLong(0);

    private SheetUpdater m_updateSheet;

    private static final Pattern FILE_EXTENSION_PATTERN = Pattern.compile(//
        Arrays.stream(ExcelFormat.values())//
            .map(ExcelFormat::getFileExtension)//
            .collect(Collectors.joining("|", "(", ")\\s*$")),
        Pattern.CASE_INSENSITIVE);

    private final ExcelTableWriterConfig m_cfg;

    private final DialogComponentStringSelection m_excelType;

    private final DialogComponentWriterFileChooser m_fileChooser;

    private final JComboBox<String>[] m_sheetNames;

    private final JLabel m_sheetNamesUpdateErr;

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

    private final DialogComponentAuthentication m_passwordComponent;

    private final SettingsModelAuthentication m_authenticationSettingsModel;

    static final String CFG_PASSWORD = "password";


    private final BiFunction<ExcelFormat, ExcelFormat, StatusMessage> m_formatErrorHandler =
            (expected, actual) -> DefaultStatusMessage.mkError("Unable to append to existing file: "
                + "you have selected \"%s\" but the existing file is actually of type \"%s\"", expected, actual);

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
                .toList());

        final SettingsModelWriterFileChooser writerModel = m_cfg.getFileChooserModel();
        final var writeFvm =
            createFlowVariableModel(writerModel.getKeysForFSLocation(), FSLocationVariableType.INSTANCE);

        m_authenticationSettingsModel = m_cfg.getAuthentication();

        final Supplier<String> passwordProvider =
            () -> CryptUtil.getPassword(m_authenticationSettingsModel, getCredentialsProvider());

        final Function<SettingsModelWriterFileChooser, StatusMessageReporter> reporter =
            fc -> StatusMessageReporterChain
                .<SettingsModelWriterFileChooser> first(new DefaultWriterStatusMessageReporter(fc))
                // no need to check if the default reporter has a problem (e.g. file does not exist)
                .successAnd(
                    // if we don't append, we do not need to check that the password can be used to decrypt
                    // the existing file
                    model -> FileOverwritePolicy.APPEND == model.getFileOverwritePolicy(),
                    new DecryptionAwareWriterStatusMessageReporter(fc, passwordProvider, m_cfg::getExcelFormat,
                        DialogUtil::decryptionErrorHandler))
                .build(fc);

        m_fileChooser = new DialogComponentWriterFileChooser(writerModel, "excel_reader_writer", writeFvm, reporter);

        m_sheetNames = Stream.generate(JComboBox<String>::new)//
            .limit(portsConfig.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length)//
            .toArray(JComboBox[]::new);
        Arrays.stream(m_sheetNames).forEach(b -> b.setEditable(true));

        m_sheetNamesUpdateErr = new JLabel();
        m_sheetNamesUpdateErr.setVisible(false);

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

        m_evaluateFormulas = new DialogComponentBoolean(m_cfg.getEvaluateFormulasModel(),
            "Evaluate formulas (leave unchecked if uncertain; see node description for details)");

        m_autoSize = new DialogComponentBoolean(m_cfg.getAutoSizeModel(), "Autosize columns");
        m_landscape = new DialogComponentButtonGroup(m_cfg.getLandscapeModel(), null, false, Orientation.values());
        m_paperSize = new JComboBox<>(PaperSize.values());
        m_openFileAfterExec =
            new DialogComponentBoolean(m_cfg.getOpenFileAfterExecModel(), "Open file after execution");
        m_openFileAfterExecLbl = new JLabel("");

        writerModel.addChangeListener(l -> toggleOptions());
        excelFormatModel.addChangeListener(l -> updateLocation());

        addTab("Settings", createSettings());

        m_passwordComponent = new DialogComponentAuthentication(m_authenticationSettingsModel, null,
            AuthenticationType.PWD, AuthenticationType.CREDENTIALS, AuthenticationType.NONE);
        m_passwordComponent.getModel().addChangeListener(e -> m_fileChooser.updateComponent());
        addTab("Encryption", createEncryptionSettingsTab());
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
        // Show decryption errors only if the file would be read (no reading on create/overwrite)
        m_sheetNamesUpdateErr.setVisible(enabled);
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

        gbc.incY().fillHorizontal();
        p.add(m_sheetNamesUpdateErr, gbc.build());
        m_fileChooser.getSettingsModel().addChangeListener(e ->
        // if we never read the file (for append), the current password does not matter
        m_sheetNamesUpdateErr
            .setVisible(m_fileChooser.getSettingsModel().getFileOverwritePolicy() == FileOverwritePolicy.APPEND));

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

    // Encryption Tab
    private JPanel createEncryptionSettingsTab() {
        final var panel = new JPanel(new GridBagLayout());
        final var gbcBuilder =
            new GBCBuilder().resetPos().anchorFirstLineStart().setWeightX(1).setWeightY(1).fillHorizontal();
        panel.add(createEncryptionPanel(), gbcBuilder.build());
        return panel;
    }

    private JPanel createEncryptionPanel() {
        final var panel = new JPanel(new GridBagLayout());
        panel.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Password to protect files"));
        final var gbcBuilder = new GBCBuilder(new Insets(0, 5, 0, 5)).resetPos().anchorFirstLineStart().fillBoth();
        panel.add(m_passwordComponent.getComponentPanel(), gbcBuilder.build());
        m_passwordComponent.getModel().addChangeListener(e -> {
            m_sheetNamesUpdateErr.setVisible(false);
            updateSheetListAndSelect();
        });
        panel.add(Box.createHorizontalBox(), gbcBuilder.incX().setWeightX(1).build());
        return panel;
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
        m_authenticationSettingsModel.saveSettingsTo(settings);
        m_passwordComponent.saveSettingsTo(settings);
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        m_excelType.loadSettingsFrom(settings, specs);
        m_fileChooser.loadSettingsFrom(settings, specs);
        m_cfg.loadSheetsInDialog(settings);
        IntStream.range(0, m_sheetNames.length)//
            .forEach(i -> {
                final var sheetNames = m_cfg.getSheetNames();
                final var name = sheetNames[i];
                m_sheetNames[i].setSelectedItem(name);
            });
        m_sheetNameCollisionHandling.loadSettingsFrom(settings, specs);
        m_writeRowKey.loadSettingsFrom(settings, specs);
        m_writeColHeader.loadSettingsFrom(settings, specs);
        try {
            m_cfg.getSkipColumnHeaderOnAppendModel().loadSettingsFrom(settings);
            m_skipColumnHeaderOnAppend.loadSettingsFrom(settings, specs);
        } catch (InvalidSettingsException e) { // NOSONAR we want to load the default here
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
        } catch (InvalidSettingsException e) { // NOSONAR we want to load the default here
            paperSizeModel.setStringValue(ExcelConstants.DEFAULT_PAPER_SIZE.name());
        }
        m_paperSize.setSelectedItem(PaperSize.valueOf(paperSizeModel.getStringValue()));
        m_openFileAfterExec.loadSettingsFrom(settings, specs);

        try {
            m_authenticationSettingsModel.loadSettingsFrom(settings);
        } catch (InvalidSettingsException e) { // NOSONAR we want to load the default here
            m_authenticationSettingsModel.setValues(AuthenticationType.NONE, null, null, null);
        }
        // Needs only be loaded
        m_passwordComponent.loadSettingsFrom(settings, specs, getCredentialsProvider());

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
            m_sheetNames[i].setModel(new DefaultComboBoxModel<>(new String[]{SCANNING})); // NOSONAR
        }

        // The id of the current update
        // Note that this code and the doneWithContext is always executed by the same thread
        // Therefore we only have to make sure that the doneWithContext belongs to the most current update
        final long currentId = m_updateSheetListId.incrementAndGet();
        if (m_updateSheet != null && !m_updateSheet.isDone()) {
            m_updateSheet.cancel(true);
        }
        m_updateSheet = new SheetUpdater(currentId);
        m_sheetNamesUpdateErr.setVisible(true);
        m_sheetNamesUpdateErr.setIcon(SharedIcons.INFO_BALLOON.get());
        m_sheetNamesUpdateErr.setText("Scanning...");
        m_updateSheet.execute();
    }

    private class SheetUpdater extends SwingWorkerWithContext<List<String>, Object> {

        final long m_ownId;

        /**
         * @param id the id of the updater which is used to decide whether to use the result.
         */
        SheetUpdater(final long id) {
            m_ownId = id;
        }

        private void logStatusMessage(final StatusMessage msg) {
            final var m = msg.getMessage();
            switch (msg.getType()) {
                case ERROR -> LOGGER.error(m);
                case INFO -> LOGGER.info(m);
                case WARNING -> LOGGER.warn(m);
            }
        }

        @Override
        protected List<String> doInBackgroundWithContext() throws Exception {
            LOGGER.debug("Refreshing sheet names...");
            try (final var accessor = m_fileChooser.getSettingsModel().createWritePathAccessor()) {
                final var path = accessor.getOutputPath(this::logStatusMessage);
                if (!Files.exists(path)) {
                    return Collections.emptyList();
                }
                final var pw = CryptUtil.getPassword(m_authenticationSettingsModel, getCredentialsProvider());
                try (final var in = FSFiles.newInputStream(path)) {
                    return ExcelUtils.readSheetNames(in, pw);
                } catch (final EncryptedDocumentException e) {
                    // provide an error message in the same style as the Excel Reader node
                    throw new EncryptedDocumentException(
                        "\"%s\" is password protected. Supply a valid password via the \"Encryption\" settings."
                            .formatted(path),
                        e);
                }
            }
        }

        @Override
        protected void doneWithContext() {
            if (m_ownId != m_updateSheetListId.get()) {
                // Another update of the sheet list has started
                // Do not update the sheet list
                return;
            }
            List<String> sheetNames = new ArrayList<>();
            final var showDecryptionError =
                m_fileChooser.getSettingsModel().getFileOverwritePolicy() == FileOverwritePolicy.APPEND;
            try {
                sheetNames = get();
                clearSheetUpdateError();
            } catch (InterruptedException e) {
                fixInterrupt();
                // ignore
            } catch (CancellationException e) { // NOSONAR: only indicates closing the dialog/newer updater
                // ignore
            } catch (ExecutionException e) { // NOSONAR we are only interested in the cause
                setSheetUpdateError(showDecryptionError, e);
            }

            for (var i = 0; i < m_sheetNames.length; i++) {
                m_sheetNames[i].setModel(new DefaultComboBoxModel<>(sheetNames.toArray(String[]::new)));
                final var selectedSheet = m_cfg.getSheetNames()[i];
                if (selectedSheet != null) {
                    m_sheetNames[i].setSelectedItem(selectedSheet);
                } else {
                    m_sheetNames[i].setSelectedIndex(!sheetNames.isEmpty() ? 0 : -1);
                }
            }
        }

        private void setSheetUpdateError(final boolean showDecryptionError, final Exception e) {
            final var cause = e.getCause();
            String statusMsg = null;
            if (cause != null) {
                statusMsg = cause.getMessage();
            }
            final var prefix = "Unable to read sheet names: ";
            m_sheetNamesUpdateErr.setIcon(SharedIcons.ERROR.get());
            m_sheetNamesUpdateErr.setText(prefix + (statusMsg == null ? "Reason unknown." : statusMsg));
            m_sheetNamesUpdateErr.setVisible(showDecryptionError);
        }

        private void clearSheetUpdateError() {
            m_sheetNamesUpdateErr.setIcon(null);
            m_sheetNamesUpdateErr.setText(null);
            m_sheetNamesUpdateErr.setVisible(false);
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
