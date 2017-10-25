package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:07
 **/
public class BooleanTypeHandler extends BaseTypeHandler<Boolean> {
    @Override
    public Boolean getValueObject(Cell cell) {
        return cell.getBooleanCellValue();
    }

    @Override
    public Cell getCell(Boolean val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
