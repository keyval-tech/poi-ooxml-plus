package com.kovizone.poi.ooxml.plus.util;


import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.annotation.Annotation;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * 属性工具类
 *
 * @author KoviChen
 */
public class ReflexUtils {

    public static Field getDeclaredField(Class<?> clazz, String fieldName) throws PoiOoxmlPlusException {
        Field field = null;
        while (!clazz.equals(Object.class)) {
            try {
                field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
                return field;
            } catch (NoSuchFieldException e) {
                clazz = clazz.getSuperclass();
            }
        }
        throw new PoiOoxmlPlusException("找不到属性：" + fieldName);
    }

    public static Annotation[] getDeclaredAnnotations(Class<?> clazz) {
        Annotation[] annotations = new Annotation[0];
        while (!clazz.equals(Object.class)) {
            Annotation[] currentAnnotations = clazz.getDeclaredAnnotations();
            for (int i = 0; i < currentAnnotations.length; i++) {
                for (Annotation annotation : annotations) {
                    if (currentAnnotations[i].annotationType().equals(annotation.annotationType())) {
                        currentAnnotations[i] = null;
                        break;
                    }
                }
            }
            annotations = ArrayUtils.addTrimAll(annotations, currentAnnotations);
            clazz = clazz.getSuperclass();
        }
        return annotations;
    }

    public static Object getValue(Object object, String fieldName) throws PoiOoxmlPlusException {
        try {
            return getValue(object, object.getClass().getDeclaredField(fieldName));
        } catch (NoSuchFieldException e) {
            throw new PoiOoxmlPlusException("找不到属性：" + fieldName);
        }
    }

    public static Object getValue(Object object, Field field) throws PoiOoxmlPlusException {
        Class<?> clazz = object.getClass();
        try {
            Method getMethod = clazz.getDeclaredMethod("get".concat(StringUtils.upperFirstCase(field.getName())));
            return getMethod.invoke(object);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
            field.setAccessible(true);
            try {
                return field.get(object);
            } catch (IllegalAccessException ex) {
                throw new PoiOoxmlPlusException("读取值失败：" + field.toString());
            }
        }
    }

    public static void setValue(Object object, String fieldName, Object value) throws PoiOoxmlPlusException {
        try {
            setValue(object, object.getClass().getDeclaredField(fieldName), value);
        } catch (NoSuchFieldException e) {
            throw new PoiOoxmlPlusException("找不到属性：" + fieldName);
        }
    }

    public static void setValue(Object object, Field field, Object value) throws PoiOoxmlPlusException {
        Class<?> clazz = object.getClass();
        try {
            Method setMethod = clazz.getDeclaredMethod("set".concat(StringUtils.upperFirstCase(field.getName())), field.getType());
            setMethod.invoke(object, value);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
            field.setAccessible(true);
            try {
                field.set(object, value);
            } catch (IllegalAccessException ex) {
                throw new PoiOoxmlPlusException("设置值失败：" + field.toString());
            }
        }
    }

    public static <T> T newInstance(Class<T> clazz) throws PoiOoxmlPlusException {
        return newInstance(clazz, new Object[0]);
    }

    @SuppressWarnings("unchecked")
    public static <T> T newInstance(Class<T> clazz, Object... params) throws PoiOoxmlPlusException {
        if (params == null) {
            params = new Object[0];
        }
        Constructor<?>[] constructors = clazz.getConstructors();
        for (Constructor<?> constructor : constructors) {
            if (constructor.getParameterTypes().length == params.length) {
                try {
                    Object entity = constructor.newInstance(params);
                    return (T) entity;
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException ignored) {
                }
            }
        }
        throw new PoiOoxmlPlusException("没有找到合适的构造方法：" + clazz.toString());
    }
}
