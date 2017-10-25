package com.github.feyond.type;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

/**
 * @author
 * @create 2017-10-11 17:34
 **/
public class ParameterizedTypeImpl implements ParameterizedType {
    private ParameterizedType delegate;
    private Type[] actualTypeArguments;

    ParameterizedTypeImpl(ParameterizedType delegate, Type[] actualTypeArguments) {
        this.delegate = delegate;
        this.actualTypeArguments = actualTypeArguments;
    }

    @Override
    public Type[] getActualTypeArguments() {
        return actualTypeArguments;
    }

    @Override
    public Type getRawType() {
        return delegate.getRawType();
    }

    @Override
    public Type getOwnerType() {
        return delegate.getOwnerType();
    }
}
