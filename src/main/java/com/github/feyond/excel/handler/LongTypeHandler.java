package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:40
 **/
public class LongTypeHandler extends BaseTypeHandler<Long> {
    @Override
    public Long getValueObject(Cell cell) {
        return (long) Math.floor(cell.getNumericCellValue());
    }

    @Override
    public Cell getCell(Long val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
