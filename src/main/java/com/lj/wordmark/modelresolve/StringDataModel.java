package com.lj.wordmark.modelresolve;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class StringDataModel implements DataModelResolve<String> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return dataModel instanceof String;
    }

    @Override
    public String resolve(String dataModel) {
        return dataModel;
    }
}
