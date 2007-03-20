package org.knime.ext.poi;

import org.eclipse.core.runtime.Plugin;
import org.knime.core.node.NodeLogger;
import org.osgi.framework.BundleContext;

/**
 * The activator class controls the plug-in life cycle.
 */
public class POIActivator extends Plugin {

    /** The plug-in ID. */
    public static final String PLUGIN_ID = "org.knime.ext.poi";

    // The shared instance
    private static POIActivator plugin;

    /**
     * The constructor.
     */
    public POIActivator() {
        plugin = this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void start(final BundleContext context) throws Exception {
        super.start(context);
        NodeLogger logger = NodeLogger.getLogger(POIActivator.class);
        logger.info("KNIME data spreadsheet plug-in started.");
        logger.info("  This plug-in uses the POI libraries from the "
                + "Jakarta POI project ");
        logger.info("  which is part of and copyrighted by the Apache "
                + "Software FOundation.");
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void stop(final BundleContext finalContext) throws Exception {
        plugin = null;
        super.stop(finalContext);
    }

    /**
     * @return the shared instance
     */
    public static POIActivator getDefault() {
        return plugin;
    }

}
