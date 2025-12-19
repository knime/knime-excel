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

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication.AuthenticationType;
import org.knime.core.webui.node.dialog.defaultdialog.setting.credentials.LegacyCredentials;
import org.knime.node.parameters.Advanced;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.layout.After;
import org.knime.node.parameters.layout.Layout;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.Effect.EffectType;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.credentials.Credentials;
import org.knime.node.parameters.widget.credentials.PasswordWidget;

/**
 * Parameters for file encryption settings in Excel files.
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH, Germany
 */
@Advanced
@Layout(EncryptionParameters.EncryptionLayout.class)
class EncryptionParameters implements NodeParameters {

    @After(ExcelFileInfoMessage.FileInfoMessageLayout.class)
    interface EncryptionLayout {
    }

    static final class EncryptionModeRef implements ParameterReference<EncryptionMode> {
    }

    static final class CredentialsRef implements ParameterReference<LegacyCredentials> {
    }

    enum EncryptionMode {
            @Label(value = "No encryption", description = "Files are not password protected.")
            NO_ENCRYPTION,

            @Label(value = "Password", description = "Files are protected with a password.")
            PASSWORD
    }

    private static final class IsPasswordEncryption implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(EncryptionModeRef.class).isOneOf(EncryptionMode.PASSWORD);
        }
    }

    @Widget(title = "File encryption",
        description = "Allows you to specify a password to decrypt files that have been protected with a "
            + "password via Excel. Files without password protection can also be read if you set a password.")
    @ValueReference(EncryptionModeRef.class)
    @ValueSwitchWidget
    EncryptionMode m_encryptionMode = EncryptionMode.NO_ENCRYPTION;

    @Widget(title = "Password", description = "The password to decrypt the Excel file.")
    @PasswordWidget
    @ValueReference(CredentialsRef.class)
    @Effect(predicate = IsPasswordEncryption.class, type = EffectType.SHOW)
    LegacyCredentials m_credentials = new LegacyCredentials(new Credentials("", ""));

    void saveToConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var authModel = excelConfig.getAuthenticationSettingsModel();

        switch (m_encryptionMode) {
            case NO_ENCRYPTION:
                authModel.setValues(AuthenticationType.NONE, null, "", "");
                break;
            case PASSWORD:
                savePasswordToAuthModel(authModel);
                break;
            default:
                throw new IllegalStateException("Unexpected encryption mode: " + m_encryptionMode);
        }
    }

    private void savePasswordToAuthModel(final SettingsModelAuthentication authModel) {
        var credentials = m_credentials.toCredentialsOrNull();
        var flowVariableName = m_credentials.getFlowVarName();
        if (flowVariableName != null) {
            authModel.setValues(AuthenticationType.CREDENTIALS, flowVariableName, "", null);
        } else if (credentials != null) {
            authModel.setValues(AuthenticationType.PWD, null, "", credentials.getPassword());
        } else {
            throw new IllegalStateException("Found neither credentials nor flow variable name.");
        }
    }

    void loadFromConfig(final ExcelMultiTableReadConfig config) {
        final var excelConfig = config.getReaderSpecificConfig();
        final var authModel = excelConfig.getAuthenticationSettingsModel();

        switch (authModel.getAuthenticationType()) {
            case NONE:
                m_encryptionMode = EncryptionMode.NO_ENCRYPTION;
                break;
            case PWD:
                m_encryptionMode = EncryptionMode.PASSWORD;
                m_credentials = new LegacyCredentials(new Credentials("", authModel.getPassword()));
                break;
            case CREDENTIALS: // legacy case
                m_encryptionMode = EncryptionMode.PASSWORD;
                m_credentials = new LegacyCredentials(authModel.getCredential());
                break;
            default:
                throw new IllegalStateException("Unexpected authentication type: " + authModel.getAuthenticationType());
        }
    }

    @Override
    public void validate() throws InvalidSettingsException {
        // validating for a non-empty password here leads to misleading error messages when copy-pasting nodes,
        // therefore no validation is done here
    }

}
