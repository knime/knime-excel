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

import org.apache.poi.ss.usermodel.PrintSetup;
import org.knime.node.parameters.widget.choices.Label;

/**
 * Enum encoding the supported paper sizes by the 'Excel table writer' node.
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
public enum PaperSize {

        /** A4 paper size. */
        @Label("A4 - 210x297 mm")
        A4_PAPERSIZE("A4 - 210x297 mm", PrintSetup.A4_PAPERSIZE),

        /** A5 paper size. */
        @Label("A5 - 148x210 mm")
        A5_PAPERSIZE("A5 - 148x210 mm", PrintSetup.A5_PAPERSIZE),

        /** Envelope 10 paper size. */
        @Label("US Envelope #10 4 1/8 x 9 1/2")
        ENVELOPE_10_PAPERSIZE("US Envelope #10 4 1/8 x 9 1/2", PrintSetup.ENVELOPE_10_PAPERSIZE),

        /** Envelope CS paper size. */
        @Label("Envelope C5 162x229 mm")
        ENVELOPE_CS_PAPERSIZE("Envelope C5 162x229 mm", PrintSetup.ENVELOPE_CS_PAPERSIZE),

        /** Envelope DL paper size. */
        @Label("Envelope DL 110x220 mm")
        ENVELOPE_DL_PAPERSIZE("Envelope DL 110x220 mm", PrintSetup.ENVELOPE_DL_PAPERSIZE),

        /** Envelope monarch paper size. */
        @Label("Envelope Monarch 98.4×190.5 mm")
        ENVELOPE_MONARCH_PAPERSIZE("Envelope Monarch 98.4×190.5 mm", PrintSetup.ENVELOPE_MONARCH_PAPERSIZE),

        /** Executive paper size. */
        @Label("US Executive 7 1/4 x 10 1/2 in")
        EXECUTIVE_PAPERSIZE("US Executive 7 1/4 x 10 1/2 in", PrintSetup.EXECUTIVE_PAPERSIZE),

        /** Legal paper size. */
        @Label("US Legal 8 1/2 x 14 in")
        LEGAL_PAPERSIZE("US Legal 8 1/2 x 14 in", PrintSetup.LEGAL_PAPERSIZE),

        /** Letter paper size. */
        @Label("US Letter 8 1/2 x 11 in")
        LETTER_PAPERSIZE("US Letter 8 1/2 x 11 in", PrintSetup.LETTER_PAPERSIZE);

    private PaperSize(final String text, final short printSetup) {
        m_text = text;
        m_printSetup = printSetup;
    }

    private final short m_printSetup;

    private final String m_text;

    /**
     * Returns the print setup associated with this paper size.
     *
     * @return the print setup associated with this paper size
     */
    public short getPrintSetup() {
        return m_printSetup;
    }

    @Override
    public String toString() {
        return m_text;
    }
}
