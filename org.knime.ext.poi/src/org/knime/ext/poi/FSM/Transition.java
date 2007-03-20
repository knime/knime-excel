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
abstract class Transition {

    /**
     * @return the state of the FSM after the transition is performed.
     */
    abstract FSMState getNextState();

    /**
     * Is called when the transition is made. Executed after the last state exit
     * action and before the new state entry action.
     * 
     * @param node the node to act on.
     */
    abstract void performAction(final Node node);

}
