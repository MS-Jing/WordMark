package com.lj.wordmark.modelresolve;

import java.util.Date;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class DateModelResolve implements DataModelResolve<Date> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return dataModel instanceof Date;
    }

    @Override
    public String resolve(Date dataModel) {
        return dataModel.toString();
    }
}
