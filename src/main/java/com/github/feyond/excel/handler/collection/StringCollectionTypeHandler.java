package com.github.feyond.excel.handler.collection;


import com.github.feyond.excel.Builder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.Arrays;
import java.util.Collection;

/**
 * @author
 * @create 2017-09-28 10:32
 **/
public class StringCollectionTypeHandler extends BaseCollectionTypeHandler<String> {

    @Override
    public Collection<String> getValueObject(Cell cell) {
        String[] arrays = StringUtils.split(cell.getStringCellValue(), DEFAULT_SPLIT_SEPARATOR);
        return Arrays.asList(arrays);
    }

    @Override
    public Cell getCell(Collection<String> val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val, DEFAULT_SPLIT_SEPARATOR));
    }
}
