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
import java.util.Arrays;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import javax.swing.BorderFactory;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
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
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;
import org.knime.filehandling.core.util.GBCBuilder;

/**
 * The dialog of the 'Excel Table Writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
final class ExcelTableWriterNodeDialog extends NodeDialogPane {

    private static final Pattern FILE_EXTENSION_PATTERN = Pattern.compile(//
        Arrays.stream(ExcelFormat.values())//
            .map(ExcelFormat::getFileExtension)//
            .collect(Collectors.joining("|", "(", ")\\s*$")),
        Pattern.CASE_INSENSITIVE);

    private final ExcelTableWriterConfig m_cfg;

    private final DialogComponentStringSelection m_excelType;

    private final DialogComponentWriterFileChooser m_fileChooser;

    private final JTextField[] m_sheetNames;

    private final DialogComponentButtonGroup m_sheetNameCollisionHandling;

    private final DialogComponentBoolean m_writeRowKey;

    private final DialogComponentBoolean m_writeColHeader;

    private final DialogComponentBoolean m_replaceMissings;

    private final DialogComponentString m_missingValPattern;

    private final DialogComponentBoolean m_evaluateFormulas;

    private final DialogComponentButtonGroup m_landscape;

    private final DialogComponentBoolean m_autoSize;

    private final JComboBox<PaperSize> m_paperSize;

    /**
     * Constructor.
     *
     * @param portsConfig the ports configuration
     */
    ExcelTableWriterNodeDialog(final PortsConfiguration portsConfig) {
        m_cfg = new ExcelTableWriterConfig(portsConfig);

        SettingsModelString excelFormatModel = m_cfg.getExcelFormatModel();
        m_excelType = new DialogComponentStringSelection(excelFormatModel, "Excel format    ",
            Arrays.stream(ExcelFormat.values())//
                .map(ExcelFormat::name)//
                .collect(Collectors.toList()));

        final SettingsModelWriterFileChooser writerModel = m_cfg.getFileChooserModel();
        final FlowVariableModel writeFvm =
            createFlowVariableModel(writerModel.getKeysForFSLocation(), FSLocationVariableType.INSTANCE);
        m_fileChooser =
            new DialogComponentWriterFileChooser(writerModel, "excel_reader_writer", writeFvm, FilterMode.FILE);

        m_sheetNames = Stream.generate(() -> new JTextField(23))//
            .limit(portsConfig.getInputPortLocation().get(ExcelTableWriterNodeFactory.SHEET_GRP_ID).length)//
            .toArray(JTextField[]::new);
        m_sheetNameCollisionHandling = new DialogComponentButtonGroup(m_cfg.getSheetExistsHandlingModel(), null, false,
            SheetNameExistsHandling.values());

        m_writeColHeader = new DialogComponentBoolean(m_cfg.getWriteColHeaderModel(), "Write column headers");
        m_writeRowKey = new DialogComponentBoolean(m_cfg.getWriteRowKeyModel(), "Write row key");
        m_replaceMissings = new DialogComponentBoolean(m_cfg.getReplaceMissingsModel(), "Replace missing values by");
        m_missingValPattern = new DialogComponentString(m_cfg.getMissingValPatternModel(), null);

        m_evaluateFormulas = new DialogComponentBoolean(m_cfg.getEvaluateFormulasModel(), "Evaluate formulas");

        m_autoSize = new DialogComponentBoolean(m_cfg.getAutoSizeModel(), "Autosize columns");
        m_landscape = new DialogComponentButtonGroup(m_cfg.getLandscapeModel(), null, false, Orientation.values());
        m_paperSize = new JComboBox<>(PaperSize.values());

        writerModel.addChangeListener(l -> toggleAppendRelatedOptions());
        excelFormatModel.addChangeListener(l -> updateLocation());

        addTab("Settings", createSettings());
    }

    private void toggleAppendRelatedOptions() {
        final boolean enabled = m_fileChooser.getSettingsModel().getFileOverwritePolicy() == FileOverwritePolicy.APPEND;
        m_evaluateFormulas.getModel().setEnabled(enabled);
        m_sheetNameCollisionHandling.getModel().setEnabled(enabled);
    }

    private void updateLocation() {
        final ExcelFormat format = ExcelFormat.valueOf(m_cfg.getExcelFormatModel().getStringValue());
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
        final JPanel p = new JPanel(new GridBagLayout());
        final GBCBuilder gbc = new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(1).fillHorizontal();
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

        gbc.incY().setWeightY(1).fillVertical();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createFileChooserPanel() {
        final JPanel p = new JPanel(new GridBagLayout());
        p.setBorder(
            BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "File format & output location"));
        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        p.add(m_excelType.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal().insetLeft(4);
        p.add(m_fileChooser.getComponentPanel(), gbc.build());

        return p;
    }

    private Component createSheetNamesPanel() {
        final JPanel p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Sheets"));
        final GBCBuilder gbc = new GBCBuilder().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone();

        for (int i = 0; i < m_sheetNames.length; i++) {
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
        final JPanel p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Names and IDs"));
        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_writeRowKey.getComponentPanel(), gbc.build());

        gbc.incY().insetTop(-5);
        p.add(m_writeColHeader.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal();
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createMissingsPanel() {
        final JPanel p = new JPanel(new GridBagLayout());
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
        final JPanel p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Formulas"));

        final GBCBuilder gbc =
            new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0).fillNone().insetLeft(-3);
        p.add(m_evaluateFormulas.getComponentPanel(), gbc.build());

        gbc.incY().setWeightX(1).fillHorizontal().insetTop(-10);
        p.add(new JPanel(), gbc.build());

        return p;
    }

    private Component createSizePanel() {
        final JPanel p = new JPanel(new GridBagLayout());
        p.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Layout"));

        final GBCBuilder gbc = new GBCBuilder().resetX().resetY().anchorLineStart().setWeightX(0).setWeightY(0)
            .fillNone().setWidth(2).insetLeft(-3);
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
            .map(JTextField::getText)//
            .map(String::trim)//
            .toArray(String[]::new));
        m_sheetNameCollisionHandling.saveSettingsTo(settings);
        m_writeRowKey.saveSettingsTo(settings);
        m_writeColHeader.saveSettingsTo(settings);
        m_replaceMissings.saveSettingsTo(settings);
        m_missingValPattern.saveSettingsTo(settings);
        m_evaluateFormulas.saveSettingsTo(settings);
        m_autoSize.saveSettingsTo(settings);
        m_landscape.saveSettingsTo(settings);
        final SettingsModelString paperSizeModel = m_cfg.getPaperSizeModel();
        paperSizeModel.setStringValue(((PaperSize)m_paperSize.getSelectedItem()).name());
        paperSizeModel.saveSettingsTo(settings);
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
        throws NotConfigurableException {
        m_excelType.loadSettingsFrom(settings, specs);
        m_fileChooser.loadSettingsFrom(settings, specs);
        m_cfg.loadSheetsInDialog(settings);
        IntStream.range(0, m_sheetNames.length)//
            .forEach(i -> m_sheetNames[i].setText(m_cfg.getSheetNames()[i]));
        m_sheetNameCollisionHandling.loadSettingsFrom(settings, specs);
        m_writeRowKey.loadSettingsFrom(settings, specs);
        m_writeColHeader.loadSettingsFrom(settings, specs);
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
        toggleAppendRelatedOptions();
        updateLocation();
    }

    @Override
    public void onClose() {
        m_fileChooser.onClose();
    }

}
