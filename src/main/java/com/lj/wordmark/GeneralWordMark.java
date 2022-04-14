package com.lj.wordmark;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author luojing
 * @Date 2022/4/13
 */
public class GeneralWordMark extends WordMark {
    public GeneralWordMark(String wordFilePath) throws IOException {
        super(wordFilePath);
    }

    public GeneralWordMark(InputStream in) throws IOException {
        super(in);
    }

    /**
     * 使用实体类来构建数据模型
     *
     * @param dataModel 数据模型的实体
     */
    public void setDataModel(Object dataModel) {
        List<Class<?>> ownOrSuperClass = getOwnOrSuperClass(dataModel);
        Map<String, String> model = new HashMap<>();
        for (Class<?> aClass : ownOrSuperClass) {
            Field[] fields = aClass.getDeclaredFields();
            for (Field field : fields) {
                try {
                    field.setAccessible(true);
                    Object data = field.get(dataModel);
                    model.put(field.getName(), data == null ? null : data.toString());
                    field.setAccessible(false);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        super.setDataModel(model);
    }

    private List<Class<?>> getOwnOrSuperClass(Object o) {
        List<Class<?>> list = new ArrayList<>();
        Class<?> aClass = o.getClass();
        while (aClass != Object.class) {
            list.add(aClass);
            aClass = aClass.getSuperclass();
        }
        return list;
    }
}
