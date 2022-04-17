package com.lj.wordmark.modelresolve;

import org.apache.poi.ss.formula.functions.T;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class NullDataModel implements DataModelResolve<Object> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return dataModel == null;
    }

    @Override
    public String resolve(Object dataModel) {
        return null;
    }
}
