/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME GmbH, Konstanz, Germany
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
 * History
 *   12 Sep 2016 (Gabor Bakos): created
 */
package org.knime.ext.poi2.node.read2;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.Collections;
import java.util.HashSet;
import java.util.SortedMap;
import java.util.TreeMap;

import org.junit.Ignore;
import org.junit.Test;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.util.Pair;

/**
 *
 * @author Gabor Bakos
 */
public class POIUtilsTest {

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#oneBasedColumnNumberChecked(java.lang.String)}.
     * @throws InvalidSettingsException The expected exception.
     */
    @Test(expected=InvalidSettingsException.class)
    public void testOneBasedColumnNumberCheckedString() throws InvalidSettingsException {
        POIUtils.oneBasedColumnNumberChecked("");
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#oneBasedColumnNumber(java.lang.String)}.
     * @throws InvalidSettingsException Should not happen.
     */
    @Test
    public void testOneBasedColumnNumberString() throws InvalidSettingsException {
        assertEquals(0, POIUtils.oneBasedColumnNumber(""));
        assertEquals(1, POIUtils.oneBasedColumnNumber("1"));
        assertEquals(26, POIUtils.oneBasedColumnNumber("26"));
        assertEquals(1, POIUtils.oneBasedColumnNumber("A"));
        assertEquals(2, POIUtils.oneBasedColumnNumber("B"));
        assertEquals(26, POIUtils.oneBasedColumnNumber("Z"));
        assertEquals(27, POIUtils.oneBasedColumnNumber("AA"));
        assertEquals(28, POIUtils.oneBasedColumnNumber("AB"));
        assertEquals(52, POIUtils.oneBasedColumnNumber("AZ"));
        assertEquals(53, POIUtils.oneBasedColumnNumber("BA"));
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#oneBasedColumnNumberChecked(int)}.
     * @throws InvalidSettingsException The expected exception.
     */
    @Test(expected = InvalidSettingsException.class)
    public void testOneBasedColumnNumberCheckedInt() throws InvalidSettingsException {
        POIUtils.oneBasedColumnNumberChecked(0);
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#oneBasedColumnNumber(int)}.
     */
    @Test
    public void testOneBasedColumnNumberInt() {
        assertNull(POIUtils.oneBasedColumnNumber(0));
        assertEquals("A", POIUtils.oneBasedColumnNumber(1));
        assertEquals("Z", POIUtils.oneBasedColumnNumber(26));
        assertEquals("AA", POIUtils.oneBasedColumnNumber(27));
        assertEquals("AB", POIUtils.oneBasedColumnNumber(28));
        assertEquals("AZ", POIUtils.oneBasedColumnNumber(52));
        assertEquals("BA", POIUtils.oneBasedColumnNumber(53));
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#isAllEmpty(java.util.SortedMap, org.knime.ext.poi2.node.read2.XLSUserSettings)}.
     */
    @Test
    public void testIsAllEmpty() {
        assertTrue(POIUtils.isAllEmpty(Collections.emptySortedMap(), null));
        XLSUserSettings settings = new XLSUserSettings();
        settings.setFirstRow0(0);
        settings.setLastRow0(-1);
        int firstRow = settings.getFirstRow0(), lastRow = settings.getLastRow0();

        SortedMap<Integer, Pair<ActualDataType, Integer>> map = new TreeMap<>();
        map.put(3, Pair.create(ActualDataType.MISSING, 3));
        assertTrue(POIUtils.isAllEmpty(map, settings));
        map.put(3, Pair.create(ActualDataType.MISSING, 7));
        assertTrue(POIUtils.isAllEmpty(map, settings));
        map.put(3, Pair.create(ActualDataType.NUMBER_INT, 3));
        assertFalse(POIUtils.isAllEmpty(map, settings));
        map.put(3, Pair.create(ActualDataType.NUMBER_INT, 8));
        assertFalse(POIUtils.isAllEmpty(map, settings));
        settings.setLastRow0(2);
        assertTrue(POIUtils.isAllEmpty(map, settings));
        settings.setLastRow0(3);
        assertFalse(POIUtils.isAllEmpty(map, settings));
        map.put(3, Pair.create(ActualDataType.ERROR, 3));
        settings.setUseErrorPattern(false);
        assertTrue(POIUtils.isAllEmpty(map, settings));
        settings.setUseErrorPattern(true);
        settings.setErrorPattern("Err");
        assertFalse(POIUtils.isAllEmpty(map, settings));
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#fillEmptyColHeaders(org.knime.ext.poi2.node.read2.XLSUserSettings, java.util.Set, java.lang.String[])}.
     */
    @Test
    @Ignore
    public void testFillEmptyColHeaders() {
        fail("Not yet implemented"); // TODO
    }

    /**
     * Test method for {@link org.knime.ext.poi2.node.read2.POIUtils#getUniqueName(java.lang.String, java.util.Set)}.
     */
    @Test
    public void testGetUniqueName() {
        HashSet<String> names = new HashSet<>();
        String firstHello = POIUtils.getUniqueName("Hello", names);
        assertEquals("Hello", firstHello);
        names.add(firstHello);
        String secondHello = POIUtils.getUniqueName("Hello", names);
        assertEquals("Hello_2", secondHello);
        names.add(secondHello);
        String firstWorld = POIUtils.getUniqueName("World", names);
        assertEquals("World", firstWorld);
        names.add(firstWorld);
        String thirdHello = POIUtils.getUniqueName("Hello", names);
        assertEquals("Hello_3", thirdHello);
    }

}
