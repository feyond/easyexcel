package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:37
 **/
public class IntegerTypeHandler extends BaseTypeHandler<Integer> {
    @Override
    public Integer getValueObject(Cell cell) {
        return (int) Math.floor(cell.getNumericCellValue());
    }

    @Override
    public Cell getCell(Integer val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
