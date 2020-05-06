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
        Long value;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                value = Long.parseLong(cell.getStringCellValue());
                break;
            case NUMERIC:
                value = (long) Math.floor(cell.getNumericCellValue());
                break;
            default:
                value = null;
                break;

        }
        return value;
    }

    @Override
    public Cell getCell(Long val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
