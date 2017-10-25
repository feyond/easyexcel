package com.github.feyond.excel;

import com.github.feyond.util.Iterables;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Date;
import java.util.List;

/**
 * @author chenfy
 * @create 2017-10-23 11:02
 **/
public class Builder {

    public static Cell createCell(Row row, int column, CellStyle style, Date value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Short value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Double value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Float value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Long value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Integer value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Byte value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(Byte.toString(value));
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, Boolean value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Cell createCell(Row row, int column, CellStyle style, String value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        cell.setCellValue(value);
        return cell;
    }

    public static Comment createComment(CreationHelper factory, Sheet sheet, Row row, Cell cell, String value) {
        ClientAnchor anchor = factory.createClientAnchor();
        Drawing drawing = sheet.createDrawingPatriarch();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 2);
        anchor.setRow1(row.getRowNum());
        anchor.setRow2(row.getRowNum() + 3);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(value));
        comment.setAuthor("厦门银商技术部");
        return comment;
    }

    public static void createTitleRow(Sheet sheet, String sheetName, CellStyle style, int rownum, int columns) {
        Row titleRow = sheet.createRow(rownum);
        titleRow.setHeightInPoints(30);
        createCell(titleRow, 0, style, sheetName);
        sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                titleRow.getRowNum(), 0, columns - 1));
    }

    public static void createHeaderRow(Sheet sheet, List<HeaderWrapper> headerList, StyleHandler styleHandler, CreationHelper creationHelper, int rownum) {
        Row headerRow = sheet.createRow(rownum);
        headerRow.setHeightInPoints(16);
        Iterables.forEach(headerList, (column, header) -> {
            Cell cell = createCell(headerRow, column, styleHandler.getHeaderStyle(column), header.getHeader());
            if (StringUtils.isNotBlank(header.getComment())) {
                Comment comment = createComment(creationHelper, sheet, headerRow, cell, header.getComment());
                cell.setCellComment(comment);
            }
        });
    }

    public static void autoSizeColumn(Sheet sheet, List<HeaderWrapper> headerList) {
        ((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
        Iterables.forEach(headerList, (column, header) -> {
            sheet.autoSizeColumn(column, true);
            int colWidth = sheet.getColumnWidth(column) * 2;
            sheet.setColumnWidth(column, colWidth < 3000 ? 3000 : colWidth);
        });
    }
}
