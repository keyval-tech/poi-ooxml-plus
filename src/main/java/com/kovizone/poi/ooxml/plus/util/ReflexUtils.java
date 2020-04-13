package com.kovizone.poi.ooxml.plus.util;


import com.kovizone.poi.ooxml.plus.exception.ReflexException;

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

    public static Field getDeclaredField(Class<?> clazz, String fieldName) throws ReflexException {
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
        throw new ReflexException("找不到属性：" + fieldName);
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

    public static Object getValue(Object object, String fieldName) throws ReflexException {
        try {
            return getValue(object, object.getClass().getDeclaredField(fieldName));
        } catch (NoSuchFieldException e) {
            throw new ReflexException("找不到属性：" + fieldName);
        }
    }

    public static Object getValue(Object object, Field field) throws ReflexException {
        Class<?> clazz = object.getClass();
        try {
            Method getMethod = clazz.getDeclaredMethod("get".concat(StringUtils.upperFirstCase(field.getName())));
            return getMethod.invoke(object);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
            field.setAccessible(true);
            try {
                return field.get(object);
            } catch (IllegalAccessException ex) {
                e.printStackTrace();
                throw new ReflexException("读取属性值失败：" + field.toString());
            }
        }
    }

    public static void setValue(Object object, String fieldName, Object value) throws ReflexException {
        try {
            setValue(object, object.getClass().getDeclaredField(fieldName), value);
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
            throw new ReflexException("找不到属性：" + fieldName);
        }
    }

    public static void setValue(Object object, Field field, Object value) throws ReflexException {
        Class<?> clazz = object.getClass();
        try {
            Method setMethod = clazz.getDeclaredMethod("set".concat(StringUtils.upperFirstCase(field.getName())), field.getType());
            setMethod.invoke(object, value);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
            field.setAccessible(true);
            try {
                field.set(object, value);
            } catch (IllegalAccessException ex) {
                e.printStackTrace();
                throw new ReflexException("设置值失败：" + field.toString());
            }
        }
    }

    public static <T> T newInstance(Class<T> clazz) throws ReflexException {
        return newInstance(clazz, new Object[0]);
    }

    @SuppressWarnings("unchecked")
    public static <T> T newInstance(Class<T> clazz, Object... params) throws ReflexException {
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
        throw new ReflexException("没有找到合适的构造方法：" + clazz.toString());
    }
}
