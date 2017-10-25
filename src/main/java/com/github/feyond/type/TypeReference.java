package com.github.feyond.type;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

/**
 * @author
 * @create 2017-10-11 17:29
 **/
public abstract class TypeReference<T> {

    private Type rawType;

    protected TypeReference() {
        Type superClass = this.getClass().getGenericSuperclass();
        if (superClass instanceof Class) {
            throw new IllegalArgumentException("Internal error: TypeReference constructed without actual type information");
        } else {
            this.rawType = ((ParameterizedType) superClass).getActualTypeArguments()[0];
        }
    }

    public final Type getRawType() {
        return this.rawType;
    }
}
