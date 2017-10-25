package com.github.feyond.excel.handler;


import com.github.feyond.excel.TypeHandler;
import com.github.feyond.type.TypeReference;

/**
 * @author
 * @create 2017-10-11 17:33
 **/
public abstract class BaseTypeHandler<T> extends TypeReference<T> implements TypeHandler<T> {
    protected static final String DEFAULT_SPLIT_SEPARATOR = ",";
}
