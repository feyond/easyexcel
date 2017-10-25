package com.github.feyond.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * 类型控制转换
 * http://www.cnblogs.com/dongying/p/4040435.html
 *
 * @author
 * @create 2017-09-25 16:59
 **/
public interface TypeHandler<T> {

    T getValueObject(Cell cell) ;

    Cell getCell(T val, CellStyle style, Row row, Integer column) ;
}
