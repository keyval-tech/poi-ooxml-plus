package com.kovizone.poi.ooxml.plus.util;

import org.apache.poi.ss.formula.functions.T;

import java.util.Arrays;

/**
 * 数据工具
 *
 * @author KoviChen
 */
public class ArrayUtils {

    /**
     * 数组叠加，去除空项
     *
     * @param original 原数组
     * @param array    叠加数组
     * @param <T>      泛型
     * @return 新数组
     */
    public static <T> T[] addTrimAll(T[] original, T[] array) {
        return trim(addAll(original, array));
    }

    /**
     * 数组叠加
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
