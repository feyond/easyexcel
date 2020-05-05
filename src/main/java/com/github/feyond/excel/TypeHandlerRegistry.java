package com.github.feyond.excel;

import com.github.feyond.excel.handler.*;
import com.github.feyond.excel.handler.array.ByteArrayTypeHandler;
import com.github.feyond.excel.handler.array.ObjectArrayTypeHandler;
import com.github.feyond.excel.handler.array.StringArrayTypeHandler;
import com.github.feyond.excel.handler.collection.*;
import com.github.feyond.util.CollectionHelper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.lang.reflect.Type;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author
 * @create 2017-09-27 15:40
 **/
@Slf4j
public class TypeHandlerRegistry {

    private final LinkedHashMap<Class<?>, TypeHandler<?>> ALL_TYPE_HANDLERS_MAP = new LinkedHashMap();
    private final LinkedHashMap<Class<?>, TypeHandler<?>> COLLECTION_TYPE_HANDLERS_MAP = new LinkedHashMap();
    private final Map<Integer, TypeHandler<?>> CACHE_COLUMN_HANDLERS_MAP = new HashMap<>();

    public TypeHandlerRegistry() {
        register(Object.class, new DefaultTypeHandler());
        register(Boolean.class, new BooleanTypeHandler());
        register(boolean.class, new BooleanTypeHandler());
        register(Short.class, new ShortTypeHandler());
        register(short.class, new ShortTypeHandler());
        register(Integer.class, new IntegerTypeHandler());
        register(int.class, new IntegerTypeHandler());
        register(Long.class, new LongTypeHandler());
        register(long.class, new LongTypeHandler());
        register(Float.class, new FloatTypeHandler());
        register(float.class, new FloatTypeHandler());
        register(Double.class, new DoubleTypeHandler());
        register(double.class, new DoubleTypeHandler());
        register(String.class, new StringTypeHandler());
        register(Date.class, new DateTypeHandler());
        register(LocalDateTime.class, new LocalDateTimeTypeHandler());

        register(Object[].class, new ObjectArrayTypeHandler());
        register(Byte[].class, new ByteArrayTypeHandler());
        register(byte[].class, new ByteArrayTypeHandler());
        register(String[].class, new StringArrayTypeHandler());

        registerCollection(Object.class, new ObjectCollectionTypeHandler());
        registerCollection(String.class, new StringCollectionTypeHandler());
        registerCollection(Integer.class, new IntegerCollectionTypeHandler());
        registerCollection(Double.class, new DoubleCollectionTypeHandler());
        registerCollection(Float.class, new FloatCollectionTypeHandler());
    }

    public TypeHandler defaultHandler() {
        return ALL_TYPE_HANDLERS_MAP.get(Object.class);
    }

    public void register(Class<?> clasz, TypeHandler typeHandler) {
        if (!ALL_TYPE_HANDLERS_MAP.containsKey(clasz)) {
            ALL_TYPE_HANDLERS_MAP.put(clasz, typeHandler);
        }
    }

    public void registerCollection(Class<?> clasz, TypeHandler typeHandler) {
        if (!COLLECTION_TYPE_HANDLERS_MAP.containsKey(clasz)) {
            COLLECTION_TYPE_HANDLERS_MAP.put(clasz, typeHandler);
        }
    }
    public void registerCollection(Class<?> type, Class<? extends TypeHandler> typeHandler) {
        try {
            if (!COLLECTION_TYPE_HANDLERS_MAP.containsKey(type)) {
                COLLECTION_TYPE_HANDLERS_MAP.put(type, typeHandler.newInstance());
            }
        } catch (Exception e) {
            log.error("注册Handler异常", e);
            throw new RuntimeException(e);
        }
    }

    public void register(Class<?> type, Class<? extends TypeHandler> hanlder) {
        try {
            if (!ALL_TYPE_HANDLERS_MAP.containsKey(type)) {
                ALL_TYPE_HANDLERS_MAP.put(type, hanlder.newInstance());
            }
        } catch (Exception e) {
            log.error("注册Handler异常", e);
            throw new RuntimeException(e);
        }
    }

    public TypeHandler getSingleTypeHandler(Type type) {
        TypeHandler<?> handler = ALL_TYPE_HANDLERS_MAP.get(type);
        if (handler == null && type != null && type instanceof Class && Enum.class.isAssignableFrom((Class<?>) type)) {
            handler = new EnumTypeHandler((Class<?>) type);
        }

        if (handler == null && type instanceof Class) {
            List<Map.Entry<Class<?>, TypeHandler<?>>> list = new LinkedList<>(ALL_TYPE_HANDLERS_MAP.entrySet());
            Collections.reverse(list);
            for (Map.Entry<Class<?>, TypeHandler<?>> entry : list) {
                if (((Class<?>) type).isAssignableFrom(entry.getKey())) {
                    handler = entry.getValue();
                    break;
                }
            }
        }
        return handler;
    }

    public TypeHandler getCollectionTypeHandler(Type rawType) {
        TypeHandler<?> handler = COLLECTION_TYPE_HANDLERS_MAP.get(rawType);

        if (handler == null && rawType != null && rawType instanceof Class && Enum.class.isAssignableFrom((Class<?>) rawType)) {
            handler = new EnumCollectionTypeHandler((Class<?>) rawType);
        }

        if (handler == null && rawType instanceof Class) {
            List<Map.Entry<Class<?>, TypeHandler<?>>> list = new LinkedList<>(COLLECTION_TYPE_HANDLERS_MAP.entrySet());
            Collections.reverse(list);
            for (Map.Entry<Class<?>, TypeHandler<?>> entry : list) {
                if (((Class<?>) rawType).isAssignableFrom(entry.getKey())) {
                    handler = entry.getValue();
                    break;
                }
            }
        }
        return handler;
    }

    public TypeHandler getTypeHandler(FieldWrapper fw, Integer column) {
        if(CACHE_COLUMN_HANDLERS_MAP.containsKey(column)) {
            return CACHE_COLUMN_HANDLERS_MAP.get(column);
        }
        boolean isCollection = Collection.class.isAssignableFrom(fw.getType());
        TypeHandler returned = null;
        if(!DefaultTypeHandler.class.isAssignableFrom(fw.getExcelField().hanlder())) {
            if(isCollection) {
                registerCollection(fw.getRawType(), fw.getExcelField().hanlder());
            } else {
                register(fw.getType(), fw.getExcelField().hanlder());
            }
        }
        if (isCollection) {
            returned = getCollectionTypeHandler(fw.getRawType());
        } else {
            returned = getSingleTypeHandler(fw.getType());
        }
        CACHE_COLUMN_HANDLERS_MAP.put(column, returned);
        return returned;
    }

    public TypeHandler getTypeHandlerByVal(Object val, Integer column) {
        if (null == val) {
            return defaultHandler();
        }
        if(CACHE_COLUMN_HANDLERS_MAP.containsKey(column)) {
            return CACHE_COLUMN_HANDLERS_MAP.get(column);
        }
        TypeHandler returned = null;
        if (val instanceof Collection) {
            Type rawType = CollectionHelper.findCommonElementType((Collection) val);
            returned = getCollectionTypeHandler(rawType);
        } else {
            returned =getSingleTypeHandler(val.getClass());
        }
        CACHE_COLUMN_HANDLERS_MAP.put(column, returned);
        return returned;
    }

    public TypeHandler getTypeHandler(Cell cell, FormulaEvaluator evaluator) {
        if(CACHE_COLUMN_HANDLERS_MAP.containsKey(cell.getColumnIndex())) {
            return CACHE_COLUMN_HANDLERS_MAP.get(cell.getColumnIndex());
        }
        if(cell.getCellTypeEnum() == CellType.FORMULA) {
            cell = evaluator.evaluateInCell(cell);
        }
        TypeHandler returned = getTypeHandler(cell, cell.getCellTypeEnum(), cell.getColumnIndex());
        CACHE_COLUMN_HANDLERS_MAP.put(cell.getColumnIndex(), returned);
        return returned;
    }

    private TypeHandler getTypeHandler(Cell cell, CellType cellTypeEnum, int columnIndex) {
        TypeHandler handler = null;
        switch (cellTypeEnum) {
            case _NONE:
                handler = ALL_TYPE_HANDLERS_MAP.get(String.class);
                break;
            case BLANK:
                handler = ALL_TYPE_HANDLERS_MAP.get(String.class);
                break;
            case STRING:
                handler = getSingleTypeHandler(String.class);
                break;
            case BOOLEAN:
                handler = getSingleTypeHandler(Boolean.class);
                break;
            case FORMULA:
                handler = getTypeHandler(cell, cell.getCachedFormulaResultTypeEnum(), columnIndex);
                break;
            case ERROR:
                handler = new ErrorTypeHandler();
                break;
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                    handler = getSingleTypeHandler(Date.class);
                } else {
                    handler = getSingleTypeHandler(Double.class);
                }
                break;
        }
        return handler;
    }
}
