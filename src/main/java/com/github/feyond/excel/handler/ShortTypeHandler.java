package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:36
 **/
public class ShortTypeHandler extends BaseTypeHandler<Short> {
    @Override
    public Short getValueObject(Cell cell) {
        return Short.parseShort(cell.getStringCellValue());
    }

    @Override
    public Cell getCell(Short val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
