package com.github.feyond.excel.handler;


import com.github.feyond.excel.Builder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;

/**
 * @author
 * @create 2017-09-27 22:45
 **/
public class LocalDateTimeTypeHandler extends BaseTypeHandler<LocalDateTime> {
    final ZoneId zoneId = ZoneId.systemDefault();
    @Override
    public LocalDateTime getValueObject(Cell cell) {
        final Date date = cell.getDateCellValue();
        return LocalDateTime.ofInstant(date.toInstant(), zoneId);
    }

    @Override
    public Cell getCell(LocalDateTime val, CellStyle style, Row row, Integer column) {
        final ZonedDateTime zdt = val.atZone(zoneId);
        return Builder.createCell(row, column, style, Date.from(zdt.toInstant()));
    }
}
