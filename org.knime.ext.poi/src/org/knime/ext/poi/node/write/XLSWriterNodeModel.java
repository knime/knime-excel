/*
 * ------------------------------------------------------------------ *
 * This source code, its documentation and all appendant files
 * are protected by copyright law. All rights reserved.
 *
 * Copyright, 2003 - 2009
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
 *   Mar 15, 2007 (ohl): created
 */
package org.knime.ext.poi.node.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;

/**
 * 
 * @author ohl, University of Konstanz
 */
public class XLSWriterNodeModel extends NodeModel {
    private XLSWriterSettings m_settings = null;

    /**
     * 
     */
    public XLSWriterNodeModel() {
        super(1, 0); // one input, no output

    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
            throws InvalidSettingsException {
        if (m_settings == null) {
            m_settings = new XLSWriterSettings();
        }
        // throws an Exception if things are not okay and sets a warning
        // message if file gets overridden.
        checkFileAccess(m_settings.getFilename());
        return new DataTableSpec[]{};
    }

    /**
     * Helper that checks some properties for the file argument.
     * 
     * @param fileName the file to check
     * @throws InvalidSettingsException if that fails
     */
    private void checkFileAccess(final String fileName)
            throws InvalidSettingsException {
        if ((fileName == null) || (fileName.length() == 0)) {
            throw new InvalidSettingsException("No output file specified.");
        }
        File file = new File(fileName);

        if (file.isDirectory()) {
            throw new InvalidSettingsException("\"" + file.getAbsolutePath()
                    + "\" is a directory.");
        }
        if (file.exists() && !m_settings.getOverwriteOK()) {
            String throwString = "File '" + file.toString()
                            + "' exists, cannot overwrite";
            throw new InvalidSettingsException(throwString);
        }
        if (!file.exists()) {
            // dunno how to check the write access to the directory. If we can't
            // create the file the execute of the node will fail. Well, too bad.
            return;
        }
        if (!file.canWrite()) {
            throw new InvalidSettingsException("Cannot write to file \""
                    + file.getAbsolutePath() + "\".");
        }
        setWarningMessage("Selected output file exists and will be "
                + "overwritten!");
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData,
            final ExecutionContext exec) throws Exception {
        File file = new File(m_settings.getFilename());
        if (file.exists() && !m_settings.getOverwriteOK()) {
            String throwString = "File '" + file.toString()
            + "' exists, cannot overwrite";
            throw new InvalidSettingsException(throwString);
        }
        FileOutputStream outStream = new FileOutputStream(file);
        
        XLSWriter xlsWriter = new XLSWriter(outStream, m_settings);

        xlsWriter.write(inData[0], exec);
        
        return new BufferedDataTable[]{};
        
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadInternals(final File nodeInternDir,
            final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        m_settings = new XLSWriterSettings(settings);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void reset() {
        // nothing to reset
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveInternals(final File nodeInternDir,
            final ExecutionMonitor exec) throws IOException,
            CanceledExecutionException {
        // nothing to save
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        if (m_settings != null) {
            m_settings.saveSettingsTo(settings);
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void validateSettings(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        String filename = new XLSWriterSettings(settings).getFilename();
        if ((filename == null) || (filename.length() == 0)) {
            throw new InvalidSettingsException("No output"
                     + " filename specified.");
        }
    }

}
