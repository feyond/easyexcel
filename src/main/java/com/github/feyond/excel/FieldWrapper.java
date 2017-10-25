package com.github.feyond.excel;

import com.github.feyond.excel.annotation.ExcelField;
import lombok.Data;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.Objects;

/**
 * Excel字段封装
 *
 * @author
 * @create 2017-09-27 9:20
 **/
@Data
public class FieldWrapper {
    private ExcelField excelField;
    private Field field;
    private Method method;
    private Class<?> type;
    private Class<?> rawType;

    public FieldWrapper(ExcelField excelField, Field field) {
        this.excelField = excelField;
        this.field = field;
        this.type = this.field.getType();
        if(this.field.getGenericType() instanceof ParameterizedType) {
            this.rawType = (Class<?>) (((ParameterizedType) this.field.getGenericType()).getActualTypeArguments()[0]);
        }
        Objects.requireNonNull(type);
    }

    public FieldWrapper(ExcelField excelField, Method method) {
        this.excelField = excelField;
        this.method = method;
        this.type = this.method.getReturnType();
        Objects.requireNonNull(type);
    }

    public FieldWrapper(ExcelField ef, Method m, OperateType type) {
        this.excelField = excelField;
        this.method = method;
        if(type == OperateType.EXPORT) {
            this.type = this.method.getReturnType();
            this.rawType = (Class<?>) (((ParameterizedType) this.method.getGenericReturnType()).getActualTypeArguments()[0]);
        }
        else {
            this.type = this.method.getParameterTypes()[0];
            this.rawType = (Class<?>) ((ParameterizedType) this.method.getGenericParameterTypes()[0]).getActualTypeArguments()[0];
        }

        Objects.requireNonNull(type);
    }

}
