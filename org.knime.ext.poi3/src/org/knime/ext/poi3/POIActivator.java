/*
 * ------------------------------------------------------------------------
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
 * -------------------------------------------------------------------
 */
package org.knime.ext.poi3;

import java.io.IOException;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.TempFile;
import org.eclipse.core.runtime.Plugin;
import org.knime.core.node.KNIMEConstants;
import org.knime.core.util.PathUtils;
import org.osgi.framework.BundleContext;

/**
 * The activator class controls the plug-in life cycle.
 */
public class POIActivator extends Plugin {

    @Override
    public void start(final BundleContext context) throws Exception {
        super.start(context);
        // disable logging of POI. To speed it up.
        System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.NullLogger");

        // We must explicitly use the standard temp dir here, otherwise it may end up in a workflow-specific temp
        // directory and be removed when the workflow is closed.
        TempFile.setTempFileCreationStrategy(new KNIMEPOITempFileCreationStrategy(
            PathUtils.createTempDir("poifiles", KNIMEConstants.getKNIMETempPath())));

        // AP-15963: lower the ratio (defaults to 0.01) as some users had issues with it
        ZipSecureFile.setMinInflateRatio(0.001d);

        // AP-16499: executing multiple excel reader/writer nodes can cause NPE during WorkbookFactory#create
        initWorkbookFactories();
    }

    private static void initWorkbookFactories() throws IOException {
        initWorkbookFactory(false);
        initWorkbookFactory(true);
    }

    private static void initWorkbookFactory(final boolean xssf) throws IOException {
        // WorkbookFactory#create initializes the appropriate factories
        try (final Workbook wb = WorkbookFactory.create(xssf)) {
            // nothing to do
        }
    }
}
