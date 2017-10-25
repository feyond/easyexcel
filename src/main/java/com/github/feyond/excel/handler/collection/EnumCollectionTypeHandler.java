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
 * @create 2017-10-23 15:43
 **/
public class EnumCollectionTypeHandler<E extends Enum<E>> extends BaseCollectionTypeHandler<E> {
    private Class<E> type;

    public EnumCollectionTypeHandler(Class<E> type) {
        if (type == null) {
            throw new IllegalArgumentException("OperateType argument cannot be null");
        } else {
            this.type = type;
        }
    }

    @Override
    public Collection<E> getValueObject(Cell cell) {
        List<E> list = new ArrayList<>();
        for (String num : StringUtils.split(cell.getStringCellValue(), DEFAULT_SPLIT_SEPARATOR)) {
            list.add(Enum.valueOf(this.type, num));
        }
        return list;
    }

    @Override
    public Cell getCell(Collection<E> val, CellStyle style, Row row, Integer column) {
        return Builder.createCell(row, column, style, StringUtils.join(val, DEFAULT_SPLIT_SEPARATOR));
    }
}
