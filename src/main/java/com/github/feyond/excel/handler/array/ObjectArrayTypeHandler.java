package com.github.feyond.excel.handler.array;

import com.github.feyond.excel.Builder;
import com.github.feyond.excel.handler.BaseTypeHandler;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author
 * @create 2017-09-29 13:22
 **/
public class ObjectArrayTypeHandler extends BaseTypeHandler<Object[]> {
    @Override
    public Object[] getValueObject(Cell cell) {
        throw new RuntimeException("Not Support");
    }

    @Override
    public Cell getCell(Object[] val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val));
    }

}
