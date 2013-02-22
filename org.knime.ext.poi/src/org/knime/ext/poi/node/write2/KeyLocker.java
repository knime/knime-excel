/*
 * ------------------------------------------------------------------------
 *
 *  Copyright (C) 2003 - 2013
 *  University of Konstanz, Germany and
 *  KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
 * Created on Feb 22, 2013 by Patrick Winter
 */
package org.knime.ext.poi.node.write2;

import java.util.HashMap;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Manages one lock per key.
 *
 * @author Patrick Winter
 */
final class KeyLocker {

    private static Lock mapLock = new ReentrantLock();

    private static HashMap<Object, CountingLock> map = new HashMap<Object, CountingLock>();

    private KeyLocker() {
        // disable default constructor
    }

    /**
     * @param key Identifies the locked object
     */
    public static void lock(final Object key) {
        mapLock.lock();
        CountingLock lock = map.get(key);
        if (lock == null) {
            lock = new CountingLock();
            map.put(key, lock);
        }
        lock.incrementReferences();
        mapLock.unlock();
        lock.lock();
    }

    /**
     * @param key Identifies the locked object
     */
    public static void unlock(final Object key) {
        mapLock.lock();
        CountingLock lock = map.get(key);
        if (lock != null) {
            lock.unlock();
            lock.decrementReferences();
            if (lock.getReferences() < 1) {
                map.remove(key);
            }
        }
        mapLock.unlock();
        if (lock == null) {
            throw new RuntimeException(key + " could not be unlocked, no lock found");
        }
    }

    private static class CountingLock {

        private Lock m_lock = new ReentrantLock();

        private int m_references = 0;

        public void lock() {
            m_lock.lock();
        }

        public void unlock() {
            m_lock.unlock();
        }

        /**
         * @return the references
         */
        public int getReferences() {
            return m_references;
        }

        public void incrementReferences() {
            ++m_references;
        }

        public void decrementReferences() {
            --m_references;
        }

    }

}
