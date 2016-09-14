/*
 * ------------------------------------------------------------------------
 *  Copyright by KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
 * -------------------------------------------------------------------
 *
 * History
 *   Apr 3, 2009 (ohl): created
 */
package org.knime.ext.poi2.node.readsheets;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.util.FileUtil;

/**
 * @author Patrick Winter, KNIME.com, Zurich, Switzerland
 */
public class XLSSheetReaderSettings {

    /**
     * Location of the selected file.
     */
    private String m_fileLocation = null;

    private static final String FILE_LOCATION = "XLS_LOCATION";

    /**
     * Saves the current values.
     *
     * @param settings Object to write values into
     */
    public void save(final NodeSettingsWO settings) {
        settings.addString(FILE_LOCATION, m_fileLocation);
    }

    /**
     * Creates a new settings object with values from the settings passed.
     *
     * @param settings the values to store in the new object
     * @return a new settings object
     * @throws InvalidSettingsException if the stored settings are incorrect.
     */
    public static XLSSheetReaderSettings load(final NodeSettingsRO settings) throws InvalidSettingsException {
        XLSSheetReaderSettings result = new XLSSheetReaderSettings();
        result.m_fileLocation = settings.getString(FILE_LOCATION);
        return result;
    }

    /**
     * Checks the settings and returns an error message - or null if everything is alright.
     *
     * @param checkFileExistence if true existence and readability of the specified file location is checked (not the
     *            content/format though).
     * @return an error message or null if settings are okay
     */
    public String getStatus(final boolean checkFileExistence) {
        if (m_fileLocation == null || m_fileLocation.isEmpty()) {
            return "No file location specified";
        }
        if (checkFileExistence) {
            try {
                URL url = new URL(m_fileLocation);
                try {
                    FileUtil.openStreamWithTimeout(url).close();
                } catch (IOException ioe) {
                    return "Can't open specified location (" + m_fileLocation + ")";
                }
            } catch (MalformedURLException mue) {
                // then try a file
                File f = new File(m_fileLocation);
                if (!f.exists()) {
                    return "Specified file doesn't exist (" + f.getAbsolutePath() + ")";
                }
                if (!f.canRead()) {
                    return "Specified file is not readable (" + f.getAbsolutePath() + ")";
                }
                if (!f.isFile()) {
                    return "Specified location is not a file (" + f.getAbsolutePath() + ")";
                }
            }
        }
        return null;
    }

    /**
     * @param fileLocation the fileLocation to set
     */
    public void setFileLocation(final String fileLocation) {
        m_fileLocation = fileLocation;
    }

    /**
     * @return the fileLocation
     */
    public String getFileLocation() {
        return m_fileLocation;
    }

}
