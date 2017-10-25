package com.github.feyond.excel.handler.array;


import com.github.feyond.excel.Builder;
import com.github.feyond.excel.handler.BaseTypeHandler;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author chenfy
 * @create 2017-10-24 9:04
 **/
public class StringArrayTypeHandler extends BaseTypeHandler<String[]> {
    @Override
    public String[] getValueObject(Cell cell) {
        return StringUtils.split(StringUtils.trimToEmpty(cell.getStringCellValue()), DEFAULT_SPLIT_SEPARATOR);
    }

    @Override
    public Cell getCell(String[] val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, new StringUtils().join(val, DEFAULT_SPLIT_SEPARATOR));
    }
}
