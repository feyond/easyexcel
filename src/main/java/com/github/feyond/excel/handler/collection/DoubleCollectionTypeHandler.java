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
 * @author
 * @create 2017-09-28 10:32
 **/
public class DoubleCollectionTypeHandler extends BaseCollectionTypeHandler<Double> {
    @Override
    public Collection<Double> getValueObject(Cell cell) {
        List<Double> list = new ArrayList<>();
        for (String num : StringUtils.split(cell.getStringCellValue(), DEFAULT_SPLIT_SEPARATOR)) {
            list.add(Double.parseDouble(num));
        }
        return list;
    }

    @Override
    public Cell getCell(Collection<Double> val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val, DEFAULT_SPLIT_SEPARATOR));
    }

}
