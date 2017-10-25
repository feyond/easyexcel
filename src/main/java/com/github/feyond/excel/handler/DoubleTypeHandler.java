package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 22:42
 **/
public class DoubleTypeHandler extends BaseTypeHandler<Double> {
    @Override
    public Double getValueObject(Cell cell) {
        return cell.getNumericCellValue();
    }

    @Override
    public Cell getCell(Double val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val);
    }
}
