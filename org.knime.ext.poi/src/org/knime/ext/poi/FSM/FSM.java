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

/**
 * 
 * @author ohl, University of Konstanz
 */
public class FSM {

    /** All events that can be performed on the states of this FSM. */
    enum Event {
        /** reset the node. */
        RESET, 
        /** execute. */
        EXECUTE, 
        /** configure the node. */
        CONFIGURE
    }

    private FSMState m_currentState;

    private final Node m_node;

    /**
     * Creates a new FSM with the specified initial state.
     * 
     * @param initialState the initial state.
     * @param node the node this FSM is abusing
     */
    public FSM(final FSMState initialState, final Node node) {
        m_currentState = initialState;
        m_node = node;
    }

    /**
     * Performs the specified event on the FSM.
     * 
     * @param e the event triggered on the current state.
     */
    public void perform(final Event e) {

        Transition t = m_currentState.getTransitionFor(e);

        // let the state know we are leaving
        m_currentState.performExitAction(t.getNextState(), m_node);
        
        // transit
        t.performAction(m_node);
        
        FSMState oldState = m_currentState;
        m_currentState = t.getNextState();
        
        // let the new state know we are coming
        m_currentState.performEntryAction(oldState, m_node);

    }

}
