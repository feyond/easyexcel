package com.github.feyond.excel.handler.array;

import com.github.feyond.excel.Builder;
import com.github.feyond.excel.handler.BaseTypeHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-28 17:13
 **/
public class ByteArrayTypeHandler extends BaseTypeHandler<byte[]> {
    @Override
    public byte[] getValueObject(Cell cell) {
        return cell.getStringCellValue().getBytes();
    }

    @Override
    public Cell getCell(byte[] val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, new String(val));
    }
}
