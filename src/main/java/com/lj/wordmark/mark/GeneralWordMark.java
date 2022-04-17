package com.lj.wordmark.mark;

import com.lj.wordmark.annotation.MarkField;
import com.lj.wordmark.modelresolve.ModelResolve;
import com.lj.wordmark.utils.LinkedMultiValueMap;
import com.lj.wordmark.utils.MultiValueMap;
import com.lj.wordmark.utils.ReflectUtils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @Author luojing
 * @Date 2022/4/15
 * D: 是基础的数据模型
 * T: 是表格中的数据模型
 * TODO 字段注解支持和嵌套引用
 */
public class GeneralWordMark<D, T> extends BaseWordMark {

    private final MultiValueMap<Class<?>, Field> classForFieldsMap = new LinkedMultiValueMap<>();

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
        Map<String, String> model = new HashMap<>();
        packageModel("", model, dataModel);
        return model;
    }

    private void packageModel(String prefixKey, Map<String, String> model, Object dataModel) {
        if (ModelResolve.ableResolve(dataModel)) {
            model.put(prefixKey, ModelResolve.resolve(dataModel));
            return;
        }
        //map直接添加
        if (dataModel instanceof Map) {
            Map dataMap = (Map) dataModel;
            dataMap.forEach((k, v) -> {
                model.put(prefixKey + k.toString(), v.toString());
            });
            return;
        }
        //不是基本数据类型继续
        List<Field> fields = getOwnAllField(dataModel.getClass());
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
                if (!ModelResolve.ableResolve(data)) {
                    key = key + ".";
                }
                packageModel(prefixKey + key, model, data);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 获取自己所有属于自己的字段
     *
     * @param aClass 类对象
     * @return 所有字段
     */
    private List<Field> getOwnAllField(Class<?> aClass) {
        List<Field> fields = classForFieldsMap.get(aClass);
        if (fields == null) {
            fields = new ArrayList<>();
            List<Class<?>> ownOrSuperClass = ReflectUtils.getOwnOrSuperClass(aClass);
            for (Class<?> superClass : ownOrSuperClass) {
                Field[] declaredFields = superClass.getDeclaredFields();
                Collections.addAll(fields, declaredFields);
            }
        }
        return fields;
    }
}
