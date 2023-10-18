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
 *   24 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import org.knime.ext.poi3.Fixtures;

/**
 * Tests for {@link CryptUtil}.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
@SuppressWarnings("static-method")
final class CryptUtilTest {


    @ParameterizedTest
    @ValueSource(strings = { Fixtures.XLSX_ENC, Fixtures.XLS_ENC, Fixtures.XLSX, Fixtures.XLS })
    void testVerifyPassword(final String filePath) throws IOException {
        try (final var in = CryptUtilTest.class.getResourceAsStream(filePath)) {
            CryptUtil.verifyPassword(in, Fixtures.TEST_PW);
        }
    }

    @ParameterizedTest
    @NullSource
    @ValueSource(strings = { Decryptor.DEFAULT_PASSWORD })
    void testVerifyPasswordDefault(final String password) throws IOException {
        try (final var in = CryptUtilTest.class.getResourceAsStream(Fixtures.XLSX_ENC_DEF)) {
            CryptUtil.verifyPassword(in, password);
        }
    }

    @ParameterizedTest
    @ValueSource(strings = { Fixtures.XLSX_ENC, Fixtures.XLS_ENC })
    void testVerifyPasswordMissing(final String filePath) throws IOException {
        try (final var in = CryptUtilTest.class.getResourceAsStream(filePath)) {
            assertThatThrownBy(() ->  CryptUtil.verifyPassword(in, null))
                .isInstanceOf(EncryptedDocumentException.class)
                .hasMessageContaining("The password is missing");
        }
    }

    @ParameterizedTest
    @ValueSource(strings = { Fixtures.XLSX_ENC, Fixtures.XLS_ENC })
    void testVerifyPasswordWrong(final String filePath) throws IOException {
        try (final var in = CryptUtilTest.class.getResourceAsStream(filePath)) {
            assertThatThrownBy(() ->  CryptUtil.verifyPassword(in, "wrong"))
                .isInstanceOf(EncryptedDocumentException.class)
                .hasMessageContaining("The given password is incorrect");
        }
    }

    @Test
    void testGetDecryptedStream() throws IOException {
        try (final var in = CryptUtilTest.class.getResourceAsStream(Fixtures.XLSX_ENC);
                final var fs = new POIFSFileSystem(in)) {
            final var root = fs.getRoot();
            try (final var decrypted = CryptUtil.getDecryptedStream(root, Fixtures.TEST_PW)) {
                assertThat(decrypted).as("decrypted stream").isNotNull();
            }
        }
    }
}
