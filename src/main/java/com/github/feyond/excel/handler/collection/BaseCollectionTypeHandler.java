package com.github.feyond.excel.handler.collection;


import com.github.feyond.excel.TypeHandler;
import com.github.feyond.type.TypeReference;

import java.util.Collection;

/**
 * @author
 * @create 2017-10-11 17:33
 **/
public abstract class BaseCollectionTypeHandler<T> extends TypeReference<T> implements TypeHandler<Collection<T>> {
    protected static final String DEFAULT_SPLIT_SEPARATOR = ",";
}
