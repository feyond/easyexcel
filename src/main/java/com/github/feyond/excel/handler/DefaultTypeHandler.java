package com.github.feyond.excel.handler;

import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 15:56
 **/
public final class DefaultTypeHandler extends BaseTypeHandler<Object> {

    @Override
    public Object getValueObject(Cell cell) {
        return cell.getStringCellValue();
    }

    @Override
    public Cell getCell(Object val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val == null ? "" : val.toString());
    }
}
