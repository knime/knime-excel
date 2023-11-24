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
 *   25 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.awt.Component;
import java.awt.Font;
import java.nio.file.Path;

import org.apache.poi.EncryptedDocumentException;
import org.knime.filehandling.core.defaultnodesettings.status.DefaultStatusMessage;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage;

/**
 * Utility for Excel dialogs.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class DialogUtil {

    private DialogUtil() {
        // hidden
    }

    /**
     * Error handler to provide status messages for decryption exceptions.
     *
     * @param path path to file
     * @param ex exception when the file cannot be decrypted
     * @return status message indicating the problem
     */
    public static StatusMessage decryptionErrorHandler(final Path path, final EncryptedDocumentException ex) {
        return DefaultStatusMessage.mkError("Unable to open password protected file \"%s\": %s. "
            + "Supply a valid password via the \"Encryption\" settings.",
            path.toString(), ex.getMessage() != null ? ex.getMessage() : "Reason unknown.");
    }

    /**
     * Italicizes the component's text by changing its font to the {@link Font#ITALIC} style.
     *
     * @param component the component of which its text should be italic
     */
    public static void italicizeText(final Component component) {
        final var font = component.getFont();
        component.setFont(new Font(font.getName(), Font.ITALIC, font.getSize()));
    }

}
