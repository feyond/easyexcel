package com.github.feyond.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @author
 * @create 2017-09-27 17:32
 **/
public class StyleHandler {

    private Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
    private final Map<Class<?>, String> ALL_TYPE_FORMAT_MAP = new HashMap();
    private SXSSFWorkbook wb;

    private static final String TITLE = "header";
    private static final String HEADER = "header";
    private static final String DATA = "data";

    public StyleHandler(SXSSFWorkbook wb) {
        this.wb = wb;

        ALL_TYPE_FORMAT_MAP.put(Boolean.class, "@");
        ALL_TYPE_FORMAT_MAP.put(boolean.class, "@");
        ALL_TYPE_FORMAT_MAP.put(Byte.class, "0");
        ALL_TYPE_FORMAT_MAP.put(byte.class, "0");
        ALL_TYPE_FORMAT_MAP.put(Short.class, "0");
        ALL_TYPE_FORMAT_MAP.put(short.class, "0");
        ALL_TYPE_FORMAT_MAP.put(Integer.class, "0");
        ALL_TYPE_FORMAT_MAP.put(int.class, "0");
        ALL_TYPE_FORMAT_MAP.put(Long.class, "0");
        ALL_TYPE_FORMAT_MAP.put(long.class, "0");
        ALL_TYPE_FORMAT_MAP.put(Float.class, "0.00");
        ALL_TYPE_FORMAT_MAP.put(float.class, "0.00");
        ALL_TYPE_FORMAT_MAP.put(Double.class, "0.00");
        ALL_TYPE_FORMAT_MAP.put(double.class, "0.00");
        ALL_TYPE_FORMAT_MAP.put(String.class, "@");
        ALL_TYPE_FORMAT_MAP.put(Date.class, "yyyy-MM-dd HH:mm:ss");
        ALL_TYPE_FORMAT_MAP.put(Character.class, "@");
        ALL_TYPE_FORMAT_MAP.put(char.class, "@");
        ALL_TYPE_FORMAT_MAP.put(Object.class, "@");

        register();
    }

    public void register() {

        registerTitle(wb);
        registerData(wb);

        CellStyle style = wb.createCellStyle();

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.LEFT);
        styles.put("data1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("data2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("data3", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put(HEADER, style);
    }

    protected void registerData(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);
    }

    protected void registerTitle(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
        style.setFont(titleFont);
        styles.put(TITLE, style);
    }

    public CellStyle getTitleStyle() {
        return styles.get(TITLE);
    }

    public CellStyle getHeaderStyle(Integer column) {
        return styles.get(HEADER);
    }

    public CellStyle getDataStyle(FieldWrapper fw, int rowNum, Integer column) {
        return getDataStyle(fw.getExcelField().format(), fw.getType(), fw.getExcelField().align(), rowNum, column);
    }

    public CellStyle getDataStyle(String format, Class aClass, int align, int rowNum, Integer column) {
        if (null == aClass) {
            return nullStyle(wb, format, align);
        }
        CellStyle style = styles.get("data_column_" + column);
        if (style == null) {
            String realFormat = StringUtils.isBlank(format) ? (StringUtils.isBlank(ALL_TYPE_FORMAT_MAP.get(aClass)) ? "@" : ALL_TYPE_FORMAT_MAP.get(aClass)) : format;
            style = wb.createCellStyle();
            style.cloneStyleFrom(styles.get("data" + (align >= 1 && align <= 3 ? align : "")));
            style.setDataFormat(wb.createDataFormat().getFormat(realFormat));
            styles.put("data_column_" + column, style);
        }
        return style;
    }

    private CellStyle nullStyle(SXSSFWorkbook wb, String format, int align) {
        CellStyle style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data" + (align >= 1 && align <= 3 ? align : "")));
        style.setDataFormat(wb.createDataFormat().getFormat(format));
        return style;
    }

}
