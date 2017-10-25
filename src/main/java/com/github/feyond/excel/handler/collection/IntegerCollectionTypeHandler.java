package com.github.feyond.excel.handler.collection;


import com.github.feyond.excel.Builder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * @author chenfy
 * @create 2017-10-23 15:11
 **/
public class IntegerCollectionTypeHandler extends BaseCollectionTypeHandler<Integer> {
    @Override
    public Collection<Integer> getValueObject(Cell cell) {
        List<Integer> list = new ArrayList<>();
        for (String num : StringUtils.split(cell.getStringCellValue(), DEFAULT_SPLIT_SEPARATOR)) {
            list.add(Integer.parseInt(num));
        }
        return list;
    }

    @Override
    public Cell getCell(Collection<Integer> val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val, DEFAULT_SPLIT_SEPARATOR));
    }
}
