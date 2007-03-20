/* ------------------------------------------------------------------
 * This source code, its documentation and all appendant files
 * are protected by copyright law. All rights reserved.
 *
 * Copyright, 2003 - 2007
 * University of Konstanz, Germany
 * Chair for Bioinformatics and Information Mining (Prof. M. Berthold)
 * and KNIME GmbH, Konstanz, Germany
 *
 * You may not modify, publish, transmit, transfer or sell, reproduce,
 * create derivative works from, distribute, perform, display, or in
 * any way exploit any of the content, in whole or in part, except as
 * otherwise expressly permitted in writing by the copyright owner or
 * as specified in the license file distributed with this product.
 *
 * If you have any questions please contact the copyright holder:
 * website: www.knime.org
 * email: contact@knime.org
 * ---------------------------------------------------------------------
 * 
 * History
 *   Mar 19, 2007 (ohl): created
 */
package org.knime.ext.poi.FSM;

import org.knime.core.node.Node;
import org.knime.ext.poi.FSM.FSM.Event;

/**
 * 
 * @author ohl, University of Konstanz
 */
abstract class FSMState {

    /** The FSM state RAW. */
    public static final FSMState RAW = FSMStateRAW.INSTANCE;

    /** The FSM state EXECUTED. */
    public static final FSMState EXECUTED = FSMStateEXECUTED.INSTANCE;

    private final String m_name;

    /**
     * Creates a new state with the specified name.
     * 
     * @param name of the state
     */
    protected FSMState(final String name) {
        if ((name == null) || (name.length() == 0)) {
            throw new IllegalArgumentException("Need a non-empty name");
        }
        m_name = name;
    }

    /**
     * Called whenever this state is going to be the new state of the FSM.
     * Executed after the transition has been performed.
     * 
     * @param fromState the previous state
     * @param node the node to abuse
     */
    abstract void performEntryAction(final FSMState fromState, final Node node);

    /**
     * Called whenever the state of the FSM is going to change and this state is
     * still the current state. Executed before the transition will perform.
     * 
     * @param toState the next state
     * @param node the node to abuse
     */
    abstract void performExitAction(final FSMState toState, final Node node);

    /**
     * Returns the transition that is to be performed when the specified event
     * occurs in _this_ state.
     * 
     * @param e the event that was triggered.
     * @return the transition if this is the current event of the FSM and the
     *         specified event is triggered.
     */
    abstract Transition getTransitionFor(final Event e);

    /**
     * {@inheritDoc}
     */
    @Override
    public String toString() {
        return m_name;
    }
}
