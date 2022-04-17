package com.lj.wordmark.modelresolve;

import com.lj.wordmark.utils.ReflectUtils;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class BaseTypeModel implements DataModelResolve<Object> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return ReflectUtils.isBaseType(dataModel.getClass());
    }

    @Override
    public String resolve(Object dataModel) {
        return dataModel.toString();
    }
}
