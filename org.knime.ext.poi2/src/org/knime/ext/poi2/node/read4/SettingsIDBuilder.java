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
 *   Nov 15, 2019 (julian): created
 */
package org.knime.ext.poi2.node.read4;

import org.knime.filehandling.core.defaultnodesettings.FileSystemChoice;
import org.knime.filehandling.core.defaultnodesettings.SettingsModelFileChooser2;

/**
 *
 * @author julian
 */
final class SettingsIDBuilder {

    private SettingsIDBuilder() {
        // Nothing to do here...
    }

    static final String getID(final SettingsModelFileChooser2 fileChooserSettings,
        final XLSUserSettings settings) {
        return "SettingsID:" + getID(fileChooserSettings) + getID(settings);
    }

    private static final String getID(final SettingsModelFileChooser2 fileChooserSettings) {
        final StringBuilder id = new StringBuilder();
        id.append(getID(fileChooserSettings.getFileOrFolderSettingsModel().getStringValue()));
        final FileSystemChoice fileSystemChoice = fileChooserSettings.getFileSystemChoice();
        id.append(getID(fileSystemChoice));
        if (FileSystemChoice.getKnimeFsChoice().equals(fileSystemChoice)
            || FileSystemChoice.getKnimeMountpointChoice().equals(fileSystemChoice)) {
            id.append(getID(fileChooserSettings.getKNIMEFileSystem()));
        } else {
            id.append("?");
        }
        return id.toString();
    }

    private static final String getID(final XLSUserSettings settings) {
        final StringBuilder id = new StringBuilder();
        id.append(getID(settings.getSheetName()));
        id.append(getID(settings.getReadAllData()));
        id.append(getID(settings.getFirstRow()));
        id.append(getID(settings.getLastRow()));
        id.append(getID(settings.getFirstColumn()));
        id.append(getID(settings.getLastColumn()));
        id.append(getID(settings.getHasColHeaders()));
        id.append(getID(settings.getColHdrRow()));
        //id of m_indexSkipJumps, m_indexContinuous are not important, those are just for generated row keys.
        id.append(getID(settings.getHasRowHeaders()));
        id.append(getID(settings.getRowHdrCol()));
        id.append(getID(settings.getMissValuePattern()));
        id.append(getID(settings.getSkipEmptyColumns()));
        id.append(getID(settings.getSkipEmptyRows()));
        id.append(getID(settings.getSkipHiddenColumns()));
        id.append(getID(settings.getKeepXLColNames()));
        id.append(getID(settings.getUniquifyRowIDs()));
        id.append(getID(settings.getUseErrorPattern()));
        id.append(getID(settings.getErrorPattern()));
        id.append(getID(settings.isReevaluateFormulae()));
        id.append(getID(settings.getTimeoutInSeconds()));
        id.append(getID(settings.isNoPreview()));
        return id.toString();
    }

    private static String getID(final String value) {
        if (value == null) {
            return "-";
        }
        if (value.isEmpty()) {
            return ".";
        }
        return value;
    }

    private static String getID(final boolean value) {
        return value ? "1" : "0";
    }

    private static String getID(final int value) {
        return "" + value;
    }

    private static String getID(final FileSystemChoice choice) {
        if (choice == null) {
            return "-";
        }
        if (choice.getId().isEmpty()) {
            return ".";
        }
        return choice.getId();
    }
}
