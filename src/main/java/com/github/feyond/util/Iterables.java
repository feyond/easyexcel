package com.github.feyond.util;

import java.util.Collection;
import java.util.Iterator;
import java.util.Objects;
import java.util.function.BiConsumer;

/**
 * @author
 * @create 2017-09-27 14:11
 **/
public class Iterables {

    public static <E> void forEach(
            Iterable<? extends E> elements, BiConsumer<Integer, ? super E> action) {
        Objects.requireNonNull(elements);
        Objects.requireNonNull(action);

        int index = 0;
        for (E element : elements) {
            action.accept(index++, element);
        }
    }

    public static int size(Iterable<?> iterable) {
        return iterable instanceof Collection ? ((Collection) iterable).size() : size(iterable.iterator());
    }

    public static int size(Iterator<?> iterator) {
        int count;
        for (count = 0; iterator.hasNext(); ++count) {
            iterator.next();
        }

        return count;
    }
}
