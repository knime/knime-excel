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
 *   Nov 13, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.util;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;

/**
 * Utility class for sheet name creation.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public final class SheetUtils {

    private SheetUtils() {
        // empty this is a utility class
    }

    /**
     * Creates a unique sheet name for the given prefix by appending the provided idx.
     *
     * @param prefix the sheet's name prefix
     * @param idx the index to append to the prefix
     * @return a unique sheet name based on the provided prefix and idx
     * @throws InvalidSettingsException - If the length of the created sheet name exceeds
     *             {@value ExcelConstants#MAX_SHEET_NAME_LENGTH}.
     */
    public static String createUniqueSheetName(final String prefix, final long idx) throws InvalidSettingsException {
        final String uniqueName = String.format("%s (%d)", prefix, idx);
        CheckUtils.checkSetting(uniqueName.length() <= ExcelConstants.MAX_SHEET_NAME_LENGTH,
            "Table spans over multiple sheets. Appending unique suffix '%s' to '%s' results in a too long sheet name "
                + "(max 31 characters). Please provide a shorter sheet name.",
            uniqueName.substring(prefix.length(), uniqueName.length()), prefix);
        return uniqueName;
    }

}
