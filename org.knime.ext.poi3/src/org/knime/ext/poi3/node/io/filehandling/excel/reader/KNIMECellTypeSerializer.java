package org.knime.ext.poi3.node.io.filehandling.excel.reader;

// TODO Header

import org.knime.base.node.io.filehandling.webui.reader2.ReaderSpecific;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell;

public interface KNIMECellTypeSerializer extends ReaderSpecific.ExternalDataTypeSerializer<ExcelCell.KNIMECellType> {

    @Override
    default String toSerializableType(ExcelCell.KNIMECellType externalType) {
        return externalType.name();
    }

    @Override
    default ExcelCell.KNIMECellType toExternalType(String serializedType)  { //            throws ReaderSpecific.ExternalDataTypeParseException {
        // TODO !!!
        return ExcelCell.KNIMECellType.valueOf(serializedType);
    }
}
