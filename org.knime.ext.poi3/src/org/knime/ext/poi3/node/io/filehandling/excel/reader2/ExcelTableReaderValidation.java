package org.knime.ext.poi3.node.io.filehandling.excel.reader2;

import org.knime.ext.poi3.node.io.filehandling.excel.writer.util.ExcelConstants;
import org.knime.node.parameters.widget.number.NumberInputWidgetValidation;

class ExcelTableReaderValidation {
    private ExcelTableReaderValidation() {
        // prevent instantiation
    }

    static class MaxNumberOfExcelRows extends NumberInputWidgetValidation.MaxValidation {

        @Override
        protected double getMax() {
            return ExcelConstants.XLSX_MAX_NUM_OF_ROWS;
        }
    }
}
