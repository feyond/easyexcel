package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-27 23:11
 **/
public class EnumTypeHandler<E extends Enum<E>> extends BaseTypeHandler<E> {
    private Class<E> type;

    public EnumTypeHandler(Class<E> type) {
        if (type == null) {
            throw new IllegalArgumentException("OperateType argument cannot be null");
        } else {
            this.type = type;
        }
    }

    @Override
    public E getValueObject(Cell cell) {
        return Enum.valueOf(this.type, cell.getStringCellValue());
    }

    @Override
    public Cell getCell(E val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, val.name());
    }
}
