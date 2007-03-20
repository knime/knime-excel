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

import org.knime.base.node.io.filereader.FileReaderNodeFactory;
import org.knime.core.node.Node;
import org.knime.ext.poi.FSM.FSM.Event;

/**
 * 
 * @author ohl, University of Konstanz
 */
public final class FSMTestClass {

    private Node m_filereader = new Node(new FileReaderNodeFactory());
    
    /**
     * 
     */
    private FSMTestClass() {
        
        FSM fsm = new FSM(FSMState.RAW, m_filereader);
        
        fsm.perform(Event.RESET);
        
        
    }

    /**
     * @param args the arguments to main
     */
    public static void main(final String[] args) {
        // huh.
    }

}
