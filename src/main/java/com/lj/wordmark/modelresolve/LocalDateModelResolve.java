package com.lj.wordmark.modelresolve;

import java.time.LocalDate;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class LocalDateModelResolve implements DataModelResolve<LocalDate> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return dataModel instanceof LocalDate;
    }

    @Override
    public String resolve(LocalDate dataModel) {
        return dataModel.toString();
    }
}
