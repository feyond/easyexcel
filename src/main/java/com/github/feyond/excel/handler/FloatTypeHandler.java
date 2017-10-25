package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:41
 **/
public class FloatTypeHandler extends BaseTypeHandler<Float> {
    @Override
    public Float getValueObject(Cell cell) {
        return new Float(cell.getNumericCellValue());
    }

    @Override
    public Cell getCell(Float val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
