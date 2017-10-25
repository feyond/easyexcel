package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.Date;

/**
 * @author
 * @create 2017-09-27 22:45
 **/
public class DateTypeHandler extends BaseTypeHandler<Date> {
    @Override
    public Date getValueObject(Cell cell) {
        return cell.getDateCellValue();
    }

    @Override
    public Cell getCell(Date val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
