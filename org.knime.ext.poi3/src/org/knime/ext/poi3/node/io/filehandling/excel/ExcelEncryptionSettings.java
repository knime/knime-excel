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
 *   11 Dec 2025 (Thomas Reifenberger): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.util.List;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModelAuthentication;
import org.knime.core.node.workflow.FlowVariable;
import org.knime.core.node.workflow.VariableType;
import org.knime.core.webui.node.dialog.defaultdialog.NodeParametersInputImpl;
import org.knime.node.parameters.NodeParameters;
import org.knime.node.parameters.NodeParametersInput;
import org.knime.node.parameters.Widget;
import org.knime.node.parameters.persistence.NodeParametersPersistor;
import org.knime.node.parameters.persistence.Persist;
import org.knime.node.parameters.persistence.Persistor;
import org.knime.node.parameters.updates.Effect;
import org.knime.node.parameters.updates.EffectPredicate;
import org.knime.node.parameters.updates.EffectPredicateProvider;
import org.knime.node.parameters.updates.ParameterReference;
import org.knime.node.parameters.updates.ValueReference;
import org.knime.node.parameters.widget.choices.ChoicesProvider;
import org.knime.node.parameters.widget.choices.FlowVariableChoicesProvider;
import org.knime.node.parameters.widget.choices.Label;
import org.knime.node.parameters.widget.choices.ValueSwitchWidget;
import org.knime.node.parameters.widget.credentials.Credentials;
import org.knime.node.parameters.widget.credentials.PasswordWidget;

/**
 *
 * Widget to manage Excel file encryption settings compatible with the legacy {@link SettingsModelAuthentication}, which
 * is used in various Excel reader and writer models. Only the parts used in those Excel nodes are supported (password &
 * a subset of the AuthTypes).
 *
 * @author Thomas Reifenberger, TNG Technology Consulting GmbH
 */
public class ExcelEncryptionSettings implements NodeParameters {

    private static final String SETTINGS_MODEL_KEY_USERNAME = "username";

    private static final String SETTINGS_MODEL_KEY_PASSWORD = "password";

    private static final String SECRET_KEY = "c-rH4Tkyk";

    /**
     * Default constructor
     */
    public ExcelEncryptionSettings() {
    }

    /**
     * All-args constructor
     *
     * @param type the authentication type
     * @param credentials Name of the workflow credentials flow variable to use if type is CREDENTIALS
     * @param password The password to use if type is PWD
     */
    public ExcelEncryptionSettings(final AuthType type, final String credentials, final Credentials password) {
        m_type = type;
        m_credentials = credentials;
        m_password = password;
    }

    /**
     * The authentication types supported for Excel file encryption.
     */
    public enum AuthType {
            /**
             * No password protection.
             */
            @Label(value = "None", description = "Only files without password protection can be updated.")
            NONE, //
            /**
             * Password protection via workflow credentials.
             */
            @Label(value = "Credentials", description = "Use a password set via workflow credentials.")
            CREDENTIALS, //
            /**
             * Password protection.
             */
            @Label(value = "Password", description = "Specify a password.")
            PWD, //
        ;
    }

    @Widget(title = "Password to protect files",
        description = "Allows you to specify a password to protect the output file with. "
            + "In case the \"append\" option is selected and the file already exists, the password must be valid "
            + "for the existing file.")
    @ValueSwitchWidget
    @Persist(configKey = "selectedType")
    @ValueReference(AuthTypeRef.class)
    AuthType m_type = AuthType.NONE;

    interface AuthTypeRef extends ParameterReference<AuthType> {
    }

    private static class IsCredentialsSelected implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(AuthTypeRef.class).isOneOf(AuthType.CREDENTIALS);
        }
    }

    private static class IsPasswordSelected implements EffectPredicateProvider {
        @Override
        public EffectPredicate init(final PredicateInitializer i) {
            return i.getEnum(AuthTypeRef.class).isOneOf(AuthType.PWD);
        }
    }

    @Widget(title = "Workflow credentials", description = "Select the workflow credentials to use.")
    @Effect(predicate = IsCredentialsSelected.class, type = Effect.EffectType.SHOW)
    @Persist(configKey = "credentials")
    @ChoicesProvider(CredentialFlowVariablesProvider.class)
    String m_credentials = "";

    @Widget(title = "Password", description = "Enter the password to protect the file.")
    @Effect(predicate = IsPasswordSelected.class, type = Effect.EffectType.SHOW)
    @Persistor(CredentialsPersistor.class)
    @PasswordWidget
    Credentials m_password = new Credentials();

    private static final class CredentialFlowVariablesProvider implements FlowVariableChoicesProvider {

        @Override
        public List<FlowVariable> flowVariableChoices(final NodeParametersInput context) {
            return context.getAvailableInputFlowVariables(VariableType.CredentialsType.INSTANCE).values().stream()
                .toList();
        }

    }

    private static final class CredentialsPersistor implements NodeParametersPersistor<Credentials> {

        @Override
        public Credentials load(final NodeSettingsRO settings) throws InvalidSettingsException {
            final var password = settings.getPassword(SETTINGS_MODEL_KEY_PASSWORD, SECRET_KEY, null);
            return new Credentials(null, password);
        }

        @Override
        public void save(final Credentials param, final NodeSettingsWO settings) {
            // username is not used, but needs to be present for compatibility with model
            settings.addString(SETTINGS_MODEL_KEY_USERNAME, null);
            settings.addPassword(SETTINGS_MODEL_KEY_PASSWORD, SECRET_KEY, param.getPassword());
        }

        @Override
        public String[][] getConfigPaths() {
            return new String[][]{{}}; // No flow variables for weakly encrypted password fields
        }

    }

    /**
     * Get the password according to the selected authentication type to be used to access referenced excel file to
     * populate the UI (e.g. loading available sheet names for a input field)
     *
     * @param context the node parameters input context
     * @return the password or null if no password is set
     */
    @SuppressWarnings("restriction")
    public String getPassword(final NodeParametersInput context) {
        if (m_type == ExcelEncryptionSettings.AuthType.NONE) {
            return null;
        } else if (m_type == ExcelEncryptionSettings.AuthType.PWD) {
            return m_password.getPassword();
        } else if (m_type == ExcelEncryptionSettings.AuthType.CREDENTIALS) {
            var flowVariableName = m_credentials;
            var credentialsProvider = ((NodeParametersInputImpl)context).getCredentialsProvider();

            try {
                if (credentialsProvider.isPresent()) {
                    return credentialsProvider.get().get(flowVariableName).getPassword();
                }
                return null;
            } catch (IllegalArgumentException e) { // Flow variable not found
                return null;
            }
        } else {
            throw new IllegalStateException("Unknown authentication type: " + m_type);
        }
    }

}
