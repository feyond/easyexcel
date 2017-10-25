package com.github.feyond.excel.handler.collection;


import com.github.feyond.excel.Builder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.Collection;

/**
 * @author
 * @create 2017-09-28 10:32
 **/
public class ObjectCollectionTypeHandler extends BaseCollectionTypeHandler<Object> {

    @Override
    public Collection<Object> getValueObject(Cell cell) {
        throw new RuntimeException("该类型控制器不支持");
    }

    @Override
    public Cell getCell(Collection<Object> val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val, DEFAULT_SPLIT_SEPARATOR));
    }
}
