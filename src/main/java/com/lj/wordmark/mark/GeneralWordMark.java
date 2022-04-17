package com.lj.wordmark.mark;

import com.lj.wordmark.annotation.MarkField;
import com.lj.wordmark.utils.LinkedMultiValueMap;
import com.lj.wordmark.utils.MultiValueMap;

import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author luojing
 * @Date 2022/4/15
 * D: 是基础的数据模型
 * T: 是表格中的数据模型
 * TODO 字段注解支持和嵌套引用
 */
public class GeneralWordMark<D, T> extends BaseWordMark {
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
    public void setDataModel(D dataModel) {
        Map<String, String> model = getObjectDataModel(dataModel);
        super.setDataModel(model);
    }

    public void setObjectTableDataModel(MultiValueMap<String, T> tableDataModel) {
        MultiValueMap<String, Map<String, String>> model = new LinkedMultiValueMap<>();
        tableDataModel.forEach((tableName, dataModels) -> {
            for (T t : dataModels) {
                Map<String, String> dataModel = getObjectDataModel(t);
                model.add(tableName, dataModel);
            }
        });
        super.setTableDataModel(model);
    }

    /**
     * 获取对象的字段组成map
     *
     * @param dataModel 数据模型对象
     * @return map
     */
    private Map<String, String> getObjectDataModel(Object dataModel) {
        List<Class<?>> ownOrSuperClass = getOwnOrSuperClass(dataModel);
        Map<String, String> model = new HashMap<>();
        for (Class<?> aClass : ownOrSuperClass) {
            //遍历属于该实体的所有字段
            Field[] fields = aClass.getDeclaredFields();
            for (Field field : fields) {
                try {
                    MarkField markField = field.getAnnotation(MarkField.class);
                    String key = null;
                    if (markField != null) {
                        key = markField.value();
                    } else {
                        key = field.getName();
                    }
                    field.setAccessible(true);
                    Object data = field.get(dataModel);
                    field.setAccessible(false);
                    model.put(key, data == null ? null : data.toString());
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        return model;
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
