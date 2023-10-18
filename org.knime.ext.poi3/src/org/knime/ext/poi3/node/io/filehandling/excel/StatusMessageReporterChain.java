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
 *   20 Oct 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.PriorityQueue;
import java.util.function.BiPredicate;
import java.util.function.Predicate;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;
import org.knime.filehandling.core.defaultnodesettings.filechooser.StatusMessageReporter;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessageUtils;

/**
 * Status message reporter chain with short-circuiting and prioritization support.
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 * @param <T> value to check status of
 */
public final class StatusMessageReporterChain<T>
        implements StatusMessageReporter {

    static final StatusMessage SUCCESS_MSG = StatusMessageUtils.SUCCESS_MSG;

    private final T m_model;

    private final Iterator<StatusMessageReporter> m_reporters;

    private final Iterator<BiPredicate<StatusMessage, T>> m_preconditions;

    private Comparator<StatusMessage> m_priorities;
    private StatusMessageReporterChain(final T model, final List<StatusMessageReporter> chain,
        final List<BiPredicate<StatusMessage, T>> preconditions, final Comparator<StatusMessage> priorities) {
        m_model = model;
        CheckUtils.checkArgument(!chain.isEmpty(), "Need at least one delegate status reporter.");
        CheckUtils.checkArgument(chain.size() != preconditions.size()  - 1,
                "Following reporters must be preceeded by precondition predicates.");
        m_reporters = chain.iterator();
        m_preconditions = preconditions.iterator();
        m_priorities = priorities;
    }

    @Override
    public StatusMessage report() throws IOException, InvalidSettingsException {
        final var status = new PriorityQueue<StatusMessage>(m_priorities);
        status.add(m_reporters.next().report());
        while (m_reporters.hasNext()) {
            final var check = m_preconditions.next();
            if (!check.test(status.peek(), m_model)) {
                return status.poll();
            }
            status.add(m_reporters.next().report());
        }
        return status.poll();
    }

    /**
     * Creates a chain with a first reporter.
     * @param <T> type of value to check
     * @param reporter first status reporter
     * @return chain builder
     */
    public static <T> Builder<T> first(final StatusMessageReporter reporter) {
        return new Builder<>(reporter);
    }

    /**
     * Builder for the status message reporter chain.
     *
     * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
     * @param <T> type of value to check
     */
    public static final class Builder<T> {

        private final List<StatusMessageReporter> m_chain = new ArrayList<>();

        private final List<BiPredicate<StatusMessage, T>> m_preconditions = new ArrayList<>();

        private final Comparator<StatusMessage> m_comp = StatusMessage::compareTo;

        private Builder(final StatusMessageReporter reporter) {
            m_chain.add(reporter);
        }

        /**
         * Appends a reporter to the chain associated with the given precondition.
         * @param precondition precondition that is checked before the reporter is tested
         * @param reporter reporter to use for test
         * @return chain builder
         */
        public Builder<T> andThen(final BiPredicate<StatusMessage, T> precondition,
            final StatusMessageReporter reporter) {
            m_preconditions.add(precondition);
            m_chain.add(reporter);
            return this;
        }

        /**
         * Appends a reporter to the chain associated with the precondition that the previous status was
         * a success message and the given precondition holds.
         *
         * @param precondition
         * @param reporter
         * @return chain builder
         */
        public Builder<T> successAnd(final Predicate<T> precondition, final StatusMessageReporter reporter) {
            return andThen((status, elem) -> {
                final var success = SUCCESS_MSG.equals(status);
                return success && precondition.test(elem);
            }, reporter);
        }

        /**
         * Appends a reporter to the chain associated with the precondition that the previous status was
         * a success message.
         *
         * @param reporter
         * @return chain builder
         */
        public Builder<T> success(final StatusMessageReporter reporter) {
            return andThen((status, elem) -> SUCCESS_MSG.equals(status), reporter);
        }


        /**
         * Builds the reporter using the given value to test on.
         *
         * @param value value to test on
         * @return status message reporter chain
         */
        public StatusMessageReporter build(final T value) {
            return new StatusMessageReporterChain<>(value, m_chain, m_preconditions, m_comp);
        }


    }
}
