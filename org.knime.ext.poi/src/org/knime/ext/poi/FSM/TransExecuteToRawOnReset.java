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
 * The transition from the EXECUTED to the RAW state on a Reset event.
 * 
 * @author ohl, University of Konstanz
 */
final class TransExecuteToRawOnReset extends Transition {

    /**
     * The singleton.
     */
    static final Transition INSTANCE = new TransExecuteToRawOnReset();

    /**
     * Don't create a new one, use the singleton above.
     */
    private TransExecuteToRawOnReset() {
        // private.
    }

    /**
     * {@inheritDoc}
     */
    @Override
    FSMState getNextState() {
        return FSMState.RAW;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    void performAction(final Node node) {
        node.resetAndConfigure();
    }

}
