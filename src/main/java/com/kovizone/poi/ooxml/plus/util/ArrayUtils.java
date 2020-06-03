package com.kovizone.poi.ooxml.plus.util;

import java.util.Arrays;

/**
 * <p>数组工具</p>
 *
 * @author KoviChen
 * @version 1.0
 */
public class ArrayUtils {

    /**
     * <p>数组叠加，去除空项</p>
     * <p><pre class="code">
     *  String[] strArray1 = new String[]{"value1", null, "value2"};
     *  String[] strArray2 = new String[]{"value3", null, "value4"};
     *
     *  // {"value1", "value2", "value3", "value4"}
     *  String[] strArray3 = ArrayUtils.addTrimAll(strArray1, strArray2);
     * </pre></p>
     *
     * @param original 原数组
     * @param array    叠加数组
     * @param <T>      泛型g
     * @return 新数组
     */
    public static <T> T[] addTrimAll(T[] original, T[] array) {
        return trim(addAll(original, array));
    }

    /**
     * <p>数组叠加</p>
     * <p>
     * <p><pre class="code">
     *  String[] strArray1 = new String[]{"value1", null, "value2"};
     *  String[] strArray2 = new String[]{"value3", null, "value4"};
     *
     *  // {"value1", null, "value2", "value3", null, "value4"}
     *  String[] strArray3 = ArrayUtils.addAll(strArray1, strArray2);
     * </pre></p>
     *
     * @param original 原数组
     * @param array    叠加数组
     * @param <T>      泛型
     * @return 新数组
     */
    public static <T> T[] addAll(T[] original, T[] array) {
        if (array == null) {
            return original;
        }
        int newLength = original.length + array.length;

        T[] newArray = Arrays.copyOf(original, newLength);

        int i = original.length;
        for (T t : array) {
            newArray[i++] = t;
        }
        return newArray;
    }

    /**
     * 去掉数组中的空项
     * <p>
     * <p><pre class="code">
     *  String[] strArray1 = new String[]{"value1", null, "value2"};
     *
     *  // {"value1", "value2"}
     *  String[] strArray2 = ArrayUtils.trim(strArray1);
     * </pre></p>
     *
     * @param original 数组
     * @param <T>      泛型
     * @return 新数组
     */
    public static <T> T[] trim(T[] original) {
        if (original == null) {
            return null;
        }
        int nullSize = 0;
        for (int i = 0; i < original.length; i++) {
            if (original[i] == null) {
                if (i < original.length - 1) {
                    for (int j = i + 1; j < original.length; j++) {
                        if (original[j] != null) {
                            original[i] = original[j];
                            original[j] = null;
                            break;
                        }
                    }
                    if (original[i] == null) {
                        nullSize++;
                    }
                } else {
                    nullSize++;
                }
            }
        }
        return Arrays.copyOf(original, original.length - nullSize);
    }
}
