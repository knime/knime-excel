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
 *   Nov 5, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.ButtonGroupEnumInterface;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;

/**
 * Enumeration of sheet selection settings.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public enum SheetSelection implements ButtonGroupEnumInterface {

        /** Select first sheet with data. */
        FIRST("First with data", SheetSelection::getFirstSheetWithData),
        /** Select sheet by name. */
        NAME("By name", SheetSelection::getSheetByName),
        /** Select sheet by index. */
        INDEX("By position", SheetSelection::getSheetByIdx);

    private final String m_text;

    private final TriFunction<Map<String, Boolean>, ExcelTableReaderConfig, Path, String> m_getSheetFunction;

    private SheetSelection(final String text,
        final TriFunction<Map<String, Boolean>, ExcelTableReaderConfig, Path, String> getSheetFunction) {
        m_text = text;
        m_getSheetFunction = getSheetFunction;
    }

    /**
     * Returns the name of the selected sheet according to the {@link SheetSelection} and configurations.
     *
     * @param sheetNames the map of sheet names to booleans that indicate if a sheet is the first with data
     * @param excelConfig the config
     * @param path the path of the file used to create meaningful error messages
     * @return the name of the selected sheet
     * @throws IOException if an I/O error occurs
     */
    public String getSelectedSheet(final Map<String, Boolean> sheetNames, final ExcelTableReaderConfig excelConfig,
        final Path path) throws IOException {
        return m_getSheetFunction.apply(sheetNames, excelConfig, path);
    }

    @Override
    public String getText() {
        return m_text;
    }

    @Override
    public String getActionCommand() {
        return name();
    }

    @Override
    public String getToolTip() {
        return null;
    }

    @Override
    public boolean isDefault() {
        return this == FIRST;
    }

    static SheetSelection loadValueInModel(final String s) throws InvalidSettingsException {
        try {
            return valueOf(s);
        } catch (IllegalArgumentException e) {
            throw new InvalidSettingsException(
                "No sheet selection setting '" + s + "' available. See node description.");
        }
    }

    private static String getSheetByIdx(final Map<String, Boolean> sheetNames, final ExcelTableReaderConfig excelConfig,
        final Path path) throws IOException {
        final int sheetIdx = excelConfig.getSheetIdx();
        if (sheetIdx < 0 || sheetIdx >= (sheetNames.size())) {
            throw new IOException("The file '" + path + "' does not contain a sheet at index " + sheetIdx + ".");
        }
        return sheetNames.keySet().toArray(new String[0])[sheetIdx];
    }

    private static String getSheetByName(final Map<String, Boolean> sheetNames,
        final ExcelTableReaderConfig excelConfig, final Path path) throws IOException {
        final String sheetName = excelConfig.getSheetName();
        if (!sheetNames.keySet().contains(sheetName)) {
            throw new IOException("The file '" + path + "' does not contain a sheet named '" + sheetName + "'.");
        }
        return sheetName;
    }

    private static String getFirstSheetWithData(final Map<String, Boolean> sheetNames,
        @SuppressWarnings("unused") final ExcelTableReaderConfig excelConfig, final Path path) throws IOException {
        if (sheetNames.isEmpty()) {
            throw new IOException("The file '" + path + "' does not contain any sheet.");
        }
        // empty should never happen; still, handle gracefully to prevent UI crashes
        return ExcelUtils.getFirstSheetWithDataOrFirstIfAllEmpty(sheetNames).orElse("");
    }

    @FunctionalInterface
    interface TriFunction<T, U, V, R> {

        R apply(T t, U u, V v) throws IOException;

        default <F> TriFunction<T, U, V, F> andThen(final Function<? super R, ? extends F> andThenFunction) {
            Objects.requireNonNull(andThenFunction);
            return (final T t, final U u, final V v) -> andThenFunction.apply(apply(t, u, v));
        }
    }
}
