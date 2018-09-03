package com.github.feyond.excel;


import com.github.feyond.util.Iterables;
import com.github.feyond.excel.annotation.ExcelField;
import com.github.feyond.excel.annotation.ExcelSheet;
import com.github.feyond.util.ReflectionHelper;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author
 * @create 2017-09-29 19:34
 **/
@Slf4j
public class AnnotationSheetWrapper extends SheetWrapper {

    /* 导出excel */
    protected List<FieldWrapper> fieldAnnotationList = new ArrayList<>();

    public AnnotationSheetWrapper(ExportExcel export, String sheetName, Class<?> cls, int... groups) {
        super(export);
        ExcelSheet es = cls.getAnnotation(ExcelSheet.class);
        initializeFields(cls, OperateType.EXPORT, groups);
        initializeSheet(StringUtils.isNotBlank(sheetName) ? sheetName : (es != null ? es.value() : DEFAULT_SHEET_NAME));
    }

    private void initializeFields(Class<?> cls, OperateType type, int... groups) {
        // Get annotation field
        Field[] fs = cls.getDeclaredFields();
        List groupList = Arrays.stream(groups).boxed().collect(Collectors.toList());
        for (Field f : fs) {
            ExcelField ef = f.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == type || ef.type() == OperateType.BOTH)) {
                List list = Arrays.stream(ef.groups()).boxed().collect(Collectors.toList());
                if (ArrayUtils.isEmpty(groups) || !Collections.disjoint(list, groupList)) {
                    fieldAnnotationList.add(new FieldWrapper(ef, f));
                }
            }
        }
        // Get annotation method
        Method[] ms = cls.getDeclaredMethods();
        for (Method m : ms) {
            ExcelField ef = m.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == type || ef.type() == OperateType.BOTH)) {
                if (ArrayUtils.isEmpty(groups) || Collections.disjoint(Arrays.asList(ef.groups()), Arrays.asList(groups))) {
                    fieldAnnotationList.add(new FieldWrapper(ef, m, type));
                }
            }
        }
        // Field sorting
        Collections.sort(fieldAnnotationList, new Comparator<FieldWrapper>() {
            @Override
            public int compare(FieldWrapper f1, FieldWrapper f2) {
                return new Integer(f1.getExcelField().sort()).compareTo(new Integer(f2.getExcelField().sort()));
            }
        });

        if (fieldAnnotationList != null) {
            fieldAnnotationList.forEach(fw -> {
                headerList.add(new HeaderWrapper(fw.getExcelField().header()));
            });
        }
    }

    /**
     * 添加数据（通过annotation.ExportField添加数据）
     *
     * @return list 数据列表
     */
    public <E> AnnotationSheetWrapper setDataList(List<E> list) {
        for (E e : list) {
            Row row = sheet.createRow(rownum++);
            log.debug("set row[{}] data... ", row.getRowNum());
            Iterables.forEach(fieldAnnotationList, (column, fw) -> {
                Object val = null;
                if (StringUtils.isNotBlank(fw.getExcelField().value())) {
                    val = ReflectionHelper.getField(e, fw.getExcelField().value());
                } else if (fw.getField() != null) {
                    if (fw.getField().isAccessible()) {
                        val = ReflectionHelper.getField(fw.getField(), e);
                    } else {
                        val = ReflectionHelper.invokeMethod(ReflectionHelper.getterMethod(e.getClass(), fw.getField().getName()), e);
                    }
                } else if (fw.getMethod() != null) {
                    val = ReflectionHelper.invokeMethod(fw.getMethod(), e);
                }
                TypeHandler handler = typeHandlerRegistry.getTypeHandler(fw, column);
                log.debug("set column[{}] data..., handler:{}", column, handler.getClass().getSimpleName());
                CellStyle style = getStyleHandler().getDataStyle(fw, row.getRowNum(), column);
                Cell cell = handler.getCell(val, style, row, column);
            });
        }
        Builder.autoSizeColumn(this.sheet, this.headerList);
        return this;
    }


    /* 读取excel */
    public AnnotationSheetWrapper(ImportExcel importExcel, int sheetIndex) {
        super(importExcel, sheetIndex);
    }

    public <E> List<E> getDataList(int dataRow, Class<E> cls, int... groups) {
        initializeFields(cls, OperateType.IMPORT, groups);
        List<E> dataList = new ArrayList<>();
        try {
            for (int rownum = dataRow; rownum <= this.sheet.getLastRowNum(); rownum++) {
                log.debug("read row[{}]...", rownum);
                E e = (E) cls.newInstance();
                Row row = sheet.getRow(rownum);
                Iterables.forEach(this.fieldAnnotationList, (column, fw) -> {
                    TypeHandler handler = typeHandlerRegistry.getTypeHandler(fw, column);
                    log.debug("read column[{}], handler:{}...", column, handler.getClass().getSimpleName());
                    Object val = handler.getValueObject(row.getCell(column));
                    if (val != null && fw.getField() != null) {
                        if (fw.getField().isAccessible()) {
                            ReflectionHelper.setField(fw.getField(), e, val);
                        } else {
                            ReflectionHelper.invokeMethod(ReflectionHelper.setterMethod(e.getClass(), fw.getField().getName(), fw.getType()), e, val);
                        }
                    } else if (val != null && fw.getMethod() != null) {
                        ReflectionHelper.invokeMethod(fw.getMethod(), e, val);
                    }
                });
                dataList.add(e);
            }
        } catch (InstantiationException | IllegalAccessException e) {
            log.error("", e);
            throw new RuntimeException("文档导入异常", e);
        }
        return dataList;
    }

}
