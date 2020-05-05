package com.github.feyond.excel;

import com.github.feyond.util.Iterables;
import com.github.feyond.util.ReflectionHelper;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.*;

/**
 * @author
 * @create 2017-09-29 19:34
 **/
@Slf4j
public class SheetWrapper {

    /* 导出Excel */
    protected static final String DEFAULT_SHEET_NAME = "Export";
    protected boolean readMergedCell;
    protected Sheet sheet;
    private ExportExcel export;
    protected TypeHandlerRegistry typeHandlerRegistry;
    private StyleHandler styleHandler;
    protected String sheetName;
    protected List<HeaderWrapper> headerList = new ArrayList<>();
    /**
     * 当前行号
     */
    protected int rownum;

    protected void initializeSheet(String sheetName) {
        this.sheetName = StringUtils.defaultIfBlank(sheetName, DEFAULT_SHEET_NAME);
        this.sheet = getWorkbook().createSheet(this.sheetName);
        if (CollectionUtils.isEmpty(headerList)) {
            throw new RuntimeException("headers must not be null!");
        }
        if (StringUtils.isNotBlank(this.sheetName) && !StringUtils.equals(DEFAULT_SHEET_NAME, this.sheetName)) {
            Builder.createTitleRow(sheet, this.sheetName, getStyleHandler().getTitleStyle(), rownum++, headerList.size());
        }
        Builder.createHeaderRow(sheet, headerList, getStyleHandler(), getWorkbook().getCreationHelper(), rownum++);
        log.debug("Sheet[{}] Initialize success.", this.sheetName);
    }


    protected StyleHandler getStyleHandler() {
        return this.styleHandler;
    }

    protected SXSSFWorkbook getWorkbook() {
        return this.export.getWb();
    }

    protected SheetWrapper(ExportExcel export) {
        this.export = export;
        this.typeHandlerRegistry = new TypeHandlerRegistry();
        this.styleHandler = new StyleHandler(getWorkbook());
    }

    public SheetWrapper(ExportExcel export, String sheetName, String[] headers) {
        this(export, sheetName, Arrays.asList(headers));
    }

    public SheetWrapper(ExportExcel export, String sheetName, List<String> headerList) {
        this(export);
        this.headerList = convertToHeaderWrapper(headerList);
        initializeSheet(sheetName);
    }

    public List<HeaderWrapper> convertToHeaderWrapper(List<String> headerList) {
        List<HeaderWrapper> list = new ArrayList<>();
        if (!CollectionUtils.isEmpty(headerList)) {
            headerList.forEach(header -> {
                list.add(new HeaderWrapper(header));
            });
        }
        return list;
    }

    public <E> SheetWrapper setDataList(List<String> fieldList, List<E> list) {
        for (E e : list) {
            Row row = sheet.createRow(rownum++);
            Iterables.forEach(fieldList, (column, field) -> {
                Object val = ReflectionHelper.getField(e, field);
                TypeHandler handler = typeHandlerRegistry.getTypeHandlerByVal(val, column);
                CellStyle style = getStyleHandler().getDataStyle(null, val == null ? null : val.getClass(), 2, row.getRowNum(), column);
                Cell cell = handler.getCell(val, style, row, column);
            });
        }
        Builder.autoSizeColumn(this.sheet, headerList);
        return this;
    }

    public SheetWrapper setDataMapList(List<String> keyList, List<Map> list) {
        for (Map e : list) {
            Row row = sheet.createRow(rownum++);
            log.debug("set row[{}] data... ", row.getRowNum());
            Iterables.forEach(keyList, (column, key) -> {
                Object val = e.get(key);
                TypeHandler handler = typeHandlerRegistry.getTypeHandlerByVal(val, column);
                log.debug("set column[{}] data..., handler:{}", column, handler.getClass().getSimpleName());
                CellStyle style = getStyleHandler().getDataStyle(null, val == null ? null : val.getClass(), 2, row.getRowNum(), column);
                Cell cell = handler.getCell(val, style, row, column);
            });
        }
        Builder.autoSizeColumn(this.sheet, headerList);
        return this;
    }

    public ExportExcel and() {
        return this.export;
    }


    /* 读取Excel */
    protected FormulaEvaluator evaluator;

    public SheetWrapper(ImportExcel importExcel, int sheetIndex) {
        this(importExcel, sheetIndex, false);
    }

    public SheetWrapper(ImportExcel importExcel, int sheetIndex, boolean readMergedCell) {
        this.typeHandlerRegistry = new TypeHandlerRegistry();
        this.sheet = importExcel.getWb().getSheetAt(sheetIndex);
        this.evaluator = importExcel.getWb().getCreationHelper().createFormulaEvaluator();
        this.readMergedCell = readMergedCell;
    }

    public List<Map<Integer,Object>> getDataList(int dataRow) {
        List<Map<Integer, Object>> dataList = new ArrayList<>();
        for (int rownum = dataRow; rownum <= this.sheet.getLastRowNum(); rownum++) {
            log.debug("read row[{}]...", rownum);
            final Map<Integer, Object> map = new HashMap<>();
            Row row = sheet.getRow(rownum);
            row.forEach(cell -> {
                TypeHandler handler =  typeHandlerRegistry.getTypeHandler(cell, evaluator);
                log.debug("read column[{}], handler:{}...", cell.getColumnIndex(), handler.getClass().getSimpleName());
                map.put(cell.getColumnIndex(), handler.getValueObject(cell));
            });
            dataList.add(map);
        }
        return dataList;
    }

    public List<Map<String,Object>> getDataListWithHeader(int headerRow) {
        List<Map<String, Object>> dataList = new ArrayList<>();
        Row headers = sheet.getRow(headerRow);
        for (int rownum = headerRow + 1; rownum <= this.sheet.getLastRowNum(); rownum++) {
            log.debug("read row[{}]...", rownum);
            final Map<String, Object> map = new HashMap<>();
            Row row = sheet.getRow(rownum);
            row.forEach(cell -> {
                if(map.containsKey(headers.getCell(cell.getColumnIndex()).getStringCellValue())) {
                    throw new RuntimeException("列标题重复");
                }
                TypeHandler handler =  typeHandlerRegistry.getTypeHandler(cell, evaluator);
                log.debug("read column[{}], handler:{}...", cell.getColumnIndex(), handler.getClass().getSimpleName());
                map.put(headers.getCell(cell.getColumnIndex()).getStringCellValue(), handler.getValueObject(cell));
            });
            dataList.add(map);
        }
        return dataList;
    }

    public List<String> getHeaders(int headerRowIndex) {
        Row headerRow = sheet.getRow(headerRowIndex);
        List<String> headers = new ArrayList<>();
        headerRow.forEach(cell -> {
            headers.add(cell.getStringCellValue());
        });
        return headers;
    }
}
