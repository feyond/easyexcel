package com.github.feyond.excel.handler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author chenfy
 * @create 2017-10-24 15:09
 **/
public class ErrorTypeHandler extends BaseTypeHandler<Byte> {

    @Override
    public Byte getValueObject(Cell cell) {
        return cell.getErrorCellValue();
    }

    @Override
    public Cell getCell(Byte val, CellStyle style, Row row, Integer column) {
        throw new RuntimeException("Not Support.");
    }

}
