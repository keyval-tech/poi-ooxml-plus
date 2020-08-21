package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.api.anno.ExcelColumn;

import java.lang.reflect.Field;
import java.util.*;

/**
 * poi-ooxml-plus 工具
 *
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class POPUtils {

    /**
     * 主要属性缓存
     */
    private static final Map<Class<?>, List<Field>> COLUMN_FIELD_LIST_CACHE = new HashMap<>(16);

    /**
     * 获取有{@link ExcelColumn}注解的属性集合，
     * 解析{@link ExcelColumn}的{@code sort}，进行排序
     *
     * @param clazz 类
     * @return 获取有@PoiColumn注解的属性集合
     */
    public static List<Field> columnFieldList(Class<?> clazz) {
        // 静态缓存
        List<Field> poiColumnFieldList = COLUMN_FIELD_LIST_CACHE.get(clazz);
        if (poiColumnFieldList == null) {
            List<Integer> sortList = new ArrayList<>(16);
            Map<Field, Integer> sortMap = new HashMap<>(16);


            while (!clazz.equals(Object.class)) {
                Field[] fields = clazz.getDeclaredFields();
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        ExcelColumn excelColumn = field.getDeclaredAnnotation(ExcelColumn.class);
                        int sort = excelColumn.sort();
                        if (!sortList.contains(sort)) {
                            sortList.add(sort);
                        }
                        sortMap.put(field, sort);
                    }
                }
                clazz = clazz.getSuperclass();
            }

            Collections.sort(sortList);
            poiColumnFieldList = new ArrayList<>();
            for (Integer sortNum : sortList) {
                Set<Map.Entry<Field, Integer>> entrySet = sortMap.entrySet();
                for (Map.Entry<Field, Integer> entry : entrySet) {
                    if (entry.getValue().equals(sortNum)) {
                        poiColumnFieldList.add(entry.getKey());
                    }
                }
            }
            COLUMN_FIELD_LIST_CACHE.put(clazz, poiColumnFieldList);
        }
        return poiColumnFieldList;
    }
}
