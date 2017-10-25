package com.github.feyond.util;

import java.util.Collection;

/**
 * @author chenfy
 * @create 2017-10-23 17:32
 **/
public abstract class CollectionHelper {

    public static Class<?> findCommonElementType(Collection<?> collection) {
        if (org.apache.commons.collections4.CollectionUtils.isEmpty(collection)) {
            return null;
        }
        Class<?> candidate = null;
        for (Object val : collection) {
            if (val != null) {
                if (candidate == null) {
                    candidate = val.getClass();
                }
                else if (candidate != val.getClass()) {
                    return null;
                }
            }
        }
        return candidate;
    }
}
