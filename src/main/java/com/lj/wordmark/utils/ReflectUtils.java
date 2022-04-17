package com.lj.wordmark.utils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class ReflectUtils {
    /**
     * 获取类自己或者父类所有的类对象
     *
     * @param aClass 目标类
     * @return 所有的类对象
     */
    public static List<Class<?>> getOwnOrSuperClass(Class<?> aClass) {
        List<Class<?>> list = new ArrayList<>();
        while (aClass != Object.class) {
            list.add(aClass);
            aClass = aClass.getSuperclass();
        }
        return list;
    }

    /**
     * 判断类对象是否是八大基本数据类型
     *
     * @param aClass 类对象
     * @return 是：true
     */
    public static boolean isBaseType(Class<?> aClass) {
        //整型 byte short int long
        return aClass == Byte.class || aClass == byte.class
                || aClass == Short.class || aClass == short.class
                || aClass == Integer.class || aClass == int.class
                || aClass == Long.class || aClass == long.class
                //浮点型 float double
                || aClass == Float.class || aClass == float.class
                || aClass == Double.class || aClass == double.class
                // 字符型 char
                || aClass == Character.class || aClass == char.class
                // 布尔 boolean
                || aClass == Boolean.class || aClass == boolean.class;
    }
}
