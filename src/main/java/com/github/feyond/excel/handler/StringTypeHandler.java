package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:43
 **/
public class StringTypeHandler extends BaseTypeHandler<String> {
    @Override
    public String getValueObject(Cell cell) {
        return cell.getStringCellValue();
    }

    @Override
    public Cell getCell(String val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
