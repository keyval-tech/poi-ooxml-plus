package com.kovizone.poi.ooxml.plus.util;

import java.util.Arrays;

/**
 * 数据工具
 *
 * @author KoviChen
 */
public class ArrayUtils {

    @SafeVarargs
    public static <T> T[] addTrimAll(T[] original, T[]... arg) {
        return trim(addAll(original, arg));
    }

    @SafeVarargs
    public static <T> T[] addAll(T[] original, T[]... arg) {
        if (arg == null) {
            return original;
        }
        int newLength = original.length;
        for (T[] array : arg) {
            if (array != null) {
                newLength += array.length;
            }
        }
        T[] newArray = Arrays.copyOf(original, newLength);
        int i = original.length;
        for (T[] array : arg) {
            if (array != null) {
                for (T t : array) {
                    newArray[i++] = t;
                }
            }
        }
        return newArray;
    }

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

    public static void main(String[] args) {
        String[] strings = new String[]{null, "3"};
        System.out.println("1: " + Arrays.toString(trim(strings)));
    }
}
