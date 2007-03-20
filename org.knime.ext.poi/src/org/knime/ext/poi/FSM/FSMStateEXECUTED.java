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
final class FSMStateEXECUTED extends FSMState {
    
    /**
     * The singleton.
     */
    static final FSMState INSTANCE = new FSMStateEXECUTED();
    
    private FSMStateEXECUTED() {
        super("EXECUTED");
    }
    
    /**
     * {@inheritDoc}
     */
    @Override
    Transition getTransitionFor(final Event e) {
        switch (e) {
        case CONFIGURE:
            return TransExecuteToRawOnReset.INSTANCE;
        case EXECUTE:
            return TransRawToExecutedOnExecute.INSTANCE;
        case RESET:
            throw new IllegalStateException("There is no transition defined "
                    + "from the state " + this + " on a RESET");
        // no default - I want the compiler to warn me if I forgot a case
        }
        return null;
    }
    
    /**
     * {@inheritDoc}
     */
    @Override
    void performEntryAction(final FSMState fromState, final Node node) {
        // notify state listeners here?
    }
    
    /**
     * {@inheritDoc}
     */
    @Override
    void performExitAction(final FSMState toState, final Node node) {
        // nothing to do.
    }



}
