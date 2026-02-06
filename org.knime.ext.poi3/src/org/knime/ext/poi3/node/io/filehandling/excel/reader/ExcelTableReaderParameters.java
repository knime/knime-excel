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
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.net.URL;

import org.knime.base.node.io.filehandling.webui.reader2.IfSchemaChangesParameters;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileReaderParameters;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionParameters;
import org.knime.base.node.io.filehandling.webui.reader2.MultiFileSelectionPath;
import org.knime.base.node.io.filehandling.webui.reader2.ReaderLayout;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.webui.node.dialog.defaultdialog.internal.widget.PersistWithin;
import org.knime.core.webui.node.dialog.defaultdialog.widget.Modification;
import org.knime.filehandling.core.node.table.reader.config.tablespec.ConfigID;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Before;
import org.knime.node.parameters.layout.Section;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;

/**
 * Excel Table Reader parameters.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@SuppressWarnings("restriction")
public class ExcelTableReaderParameters implements NodeParameters {

    @Section(title = "File and Sheet")
    @Before(ReaderLayout.DataArea.class)
    interface FileAndSheetSection {
        interface Source {
        }
    }

    @Section(title = "Values")
    @After(ReaderLayout.DataArea.class)
    @Before(ReaderLayout.ColumnAndDataTypeDetection.class)
    @Advanced
    interface ValuesSection {

    }

    ExcelTableReaderParameters(final URL url) {
        m_multiFileSelectionParams = new MultiFileSelectionParameters(url);
    }

    ExcelTableReaderParameters() {
        // default constructor
    }

    interface EncryptionParametersRef extends ParameterReference<EncryptionParameters> {
    }

    ConfigID saveToConfig(final ExcelMultiTableReadConfig config) {
        final var tableReadConfig = config.getTableReadConfig();
        tableReadConfig.setDecorateRead(false);
        tableReadConfig.setAllowShortRows(true);

        m_ifSchemaChangesParams.saveToConfig(config);
        m_multiFileReaderParams.saveToConfig(config);

        m_selectSheetParams.saveToConfig(config);
        m_readAreaParams.saveToConfig(config);
        m_columnNameParams.saveToConfig(config);
        m_rowIdParams.saveToConfig(config);
        m_skipParams.saveToConfig(config);
        m_schemaDetectionParams.saveToConfig(config);
        m_valuesParams.saveToConfig(config);
        m_encryptionParams.saveToConfig(config);

        return config.getConfigID();
    }

    @Override
    public void validate() throws InvalidSettingsException {
        m_columnNameParams.validate();
        m_encryptionParams.validate();
        m_excelFileInfoMessageParams.validate();
        m_ifSchemaChangesParams.validate();
        m_multiFileReaderParams.validate();
        m_multiFileSelectionParams.validate();
        m_readAreaParams.validate();
        m_rowIdParams.validate();
        m_schemaDetectionParams.validate();
        m_selectSheetParams.validate();
        m_skipParams.validate();
        m_valuesParams.validate();
    }

    void saveToSource(final MultiFileSelectionPath sourceSettings) {
        m_multiFileSelectionParams.saveToSource(sourceSettings);
    }

    // Common parameters
    static final class SetExcelTableReaderExtensions
        extends MultiFileSelectionParameters.SetFileReaderWidgetExtensions {
        @Override
        protected String[] getExtensions() {
            return new String[]{"xlsx", "xls", "xlsm", "xlsb"};
        }
    }

    static final class OverrideExcelReaderSourceLayout extends MultiFileSelectionParameters.OverrideReaderSourceLayout {
        @Override
        protected Class<?> getSourceLayoutClass() {
            return FileAndSheetSection.class;
        }
    }

    @Modification({SetExcelTableReaderExtensions.class, OverrideExcelReaderSourceLayout.class})
    MultiFileSelectionParameters m_multiFileSelectionParams = new MultiFileSelectionParameters();

    IfSchemaChangesParameters m_ifSchemaChangesParams = new IfSchemaChangesParameters();

    MultiFileReaderParameters m_multiFileReaderParams = new MultiFileReaderParameters();

    // ExcelTableReader-specific parameters

    @PersistWithin.PersistEmbedded // skip persisting an empty entry, this field only contains one Void field
    ExcelFileInfoMessage m_excelFileInfoMessageParams = new ExcelFileInfoMessage();

    SelectSheetParameters m_selectSheetParams = new SelectSheetParameters();

    ReadAreaParameters m_readAreaParams = new ReadAreaParameters();

    ColumnNameParameters m_columnNameParams = new ColumnNameParameters();

    RowIdParameters m_rowIdParams = new RowIdParameters();

    SkipParameters m_skipParams = new SkipParameters();

    SchemaDetectionParameters m_schemaDetectionParams = new SchemaDetectionParameters();

    ValuesParameters m_valuesParams = new ValuesParameters();

    @ValueReference(EncryptionParametersRef.class)
    EncryptionParameters m_encryptionParams = new EncryptionParameters();

    String getSourcePath() {
        return m_multiFileSelectionParams.getSourcePath();
    }

    MultiFileReaderParameters getMultiFileParameters() {
        return m_multiFileReaderParams;
    }

    IfSchemaChangesParameters getIfSchemaChangesParameters() {
        return m_ifSchemaChangesParams;
    }

}
