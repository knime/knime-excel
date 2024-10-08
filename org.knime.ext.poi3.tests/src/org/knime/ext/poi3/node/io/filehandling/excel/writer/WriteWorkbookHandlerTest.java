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
 *   8 Oct 2024 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.writer;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.Consumer;

import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.knime.ext.poi3.Fixtures;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.ExcelTableWriterNodeModel.WriteWorkbookHandler;
import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelFormat;

/**
 * Tests for WriteWorkbookHandler.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public final class WriteWorkbookHandlerTest {

    private static final String TEST_AUTHOR_UNENCRYPTED_XLSX = "KNIME Testauthor";

    @SuppressWarnings("static-method")
    @Test
    void testCreateNewWorkbook(@TempDir final Path tempPath) throws IOException, InvalidFormatException {
        // path would specify existing file, but we want to create a new one
        final var file = tempPath.resolve("testWorkbook.xlsx");

        final var factory = ExcelFormat.XLSX;
        try (final var handler = new WriteWorkbookHandler(factory, null, false, null);
                final var wb = handler.getWorkbook()) {
            handler.saveFile(file);
        }
        withCoreProperties(file.toFile(), props -> {
            assertThat(props.getCreator()) //
                .as("Newly KNIME-created \"*.%s\" files should not have creator metadata set", factory) //
                .isNull();
            assertThat(props.getLastModifiedByUser()) //
                .as("Newly KNIME-created \"*.%s\" files should not have last-modified-user metadata set", factory) //
                .isNull();
        });
    }

    @SuppressWarnings("static-method")
    @Test
    void testAppendExistingWorkbook(@TempDir final Path tempPath) throws IOException, InvalidFormatException {
        final var file = tempPath.resolve("testWorkbook.xlsx");
        try (final var in = WriteWorkbookHandlerTest.class.getResourceAsStream(Fixtures.XLSX)) {
            Files.copy(in, file);
        }

        final var factory = ExcelFormat.XLSX;
        try (final var handler = new WriteWorkbookHandler(factory, file, false, null);
                final var wb = handler.getWorkbook()) {
            handler.saveFile(file);
        }
        withCoreProperties(file.toFile(), props -> {
            assertThat(props.getCreator()) //
                .as("Existing \"*.%s\" files should not have their author overwritten or removed", factory)
                .isEqualTo(TEST_AUTHOR_UNENCRYPTED_XLSX);
            // not checking the last-modified-user here, since that is not as important
            // and apparently not possible to modify with POI in existing XLSX files
        });
    }

    /**
     * Access to the core properties of an XLSX file.
     *
     * @param file file to open
     * @param corePropsConsumer consumer for core properties
     * @throws InvalidFormatException in case the file is not a valid XLSX file
     * @throws IOException in case the file cannot be read
     */
    public static void withCoreProperties(final File file, final Consumer<CoreProperties> corePropsConsumer)
        throws IOException, InvalidFormatException {
        // Don't use WorkbookFactory, it uses ServiceLoader and provides classes from a different classloader :(
        // In that case ClassCastException will appear when trying to cast to XSSFWorkbook
        try (final var wb = new XSSFWorkbook(file)) {
            final var props = wb.getProperties().getCoreProperties();
            corePropsConsumer.accept(props);
        }
    }

}
