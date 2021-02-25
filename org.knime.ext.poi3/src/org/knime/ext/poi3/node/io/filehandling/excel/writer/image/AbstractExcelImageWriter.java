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
 *   Nov 10, 2020 (Mark Ortmann, KNIME GmbH, Berlin, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer.image;

import java.awt.Dimension;
import java.util.Arrays;
import java.util.OptionalInt;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.image.png.PNGImageContent;
import org.knime.core.data.image.png.PNGImageValue;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;

/**
 *
 * @author Mark Ortmann, KNIME GmbH, Berlin, Germany
 */
abstract class AbstractExcelImageWriter implements ExcelImageWriter {

    private final int[] m_pngIndices;

    private final int[] m_colWidth;

    /**
     * @param spec
     */
    AbstractExcelImageWriter(final DataTableSpec spec) {
        m_pngIndices = IntStream.range(0, spec.getNumColumns())//
            .filter(i -> spec.getColumnSpec(i).getType().isCompatible(PNGImageValue.class))//
            .toArray();
        m_colWidth = new int[m_pngIndices.length];
        reset();
    }

    @Override
    public final OptionalInt writeImages(final Sheet sheet, final DataRow row, final int rowIdx,
        final int colIdxOffset) {
        return IntStream.range(0, m_pngIndices.length)//
            .map(idx -> writeImages(sheet, row, rowIdx, idx, colIdxOffset))//
            .max();
    }

    private int writeImages(final Sheet sheet, final DataRow row, final int rowIdx, final int idx,
        final int colIdxOffset) {
        final int colIdx = m_pngIndices[idx];
        if (row.getCell(colIdx).isMissing()) {
            return 0;
        } else {
            final PNGImageContent image = ((PNGImageValue)row.getCell(colIdx)).getImageContent();
            final AnchorInfo anchorInfo = getAnchorInfo(image.getPreferredSize());
            createPicture(sheet, rowIdx, colIdx + colIdxOffset, image, anchorInfo);
            m_colWidth[idx] = Math.max(m_colWidth[idx], anchorInfo.m_widthInPixel);
            return anchorInfo.m_heightInPixel;
        }
    }

    private static void createPicture(final Sheet sheet, final int rowIdx, final int colIdx,
        final PNGImageContent image, final AnchorInfo anchorInfo) {
        @SuppressWarnings("resource")
        final int pictureIdx = sheet.getWorkbook().addPicture(image.getByteArray(), Workbook.PICTURE_TYPE_PNG);

        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        final ClientAnchor anchor =
            drawing.createAnchor(0, 0, anchorInfo.m_dx, anchorInfo.m_dy, colIdx, rowIdx, colIdx, rowIdx);
        anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
        drawing.createPicture(anchor, pictureIdx);
    }

    @Override
    public final void setColWidth(final Sheet sheet, final int colIdxOffset) {
        IntStream.range(0, m_pngIndices.length)//
            .filter(i -> m_colWidth[i] > -1)//
            .forEach(i -> sheet.setColumnWidth(m_pngIndices[i] + colIdxOffset, getWidth(i)));
    }

    @Override
    public final void reset() {
        Arrays.fill(m_colWidth, -1);
    }

    protected abstract AnchorInfo getAnchorInfo(Dimension dimension);

    private int getWidth(final int i) {
        return Math.min((int)(m_colWidth[i] * ExcelConstants.PIXEL_TO_COL_WIDTH), ExcelConstants.MAX_COL_WIDTH);
    }

    protected static final class AnchorInfo {

        private final int m_dy;

        private final int m_dx;

        private final int m_widthInPixel;

        private final int m_heightInPixel;

        protected AnchorInfo(final int dx, final int dy, final int widthInPixel, final int heightInPixel) {
            m_dx = dx;
            m_dy = dy;
            m_widthInPixel = widthInPixel;
            m_heightInPixel = heightInPixel;
        }

    }
}
