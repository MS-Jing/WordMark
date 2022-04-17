package com.lj.wordmark.modelresolve;

import java.time.LocalDateTime;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class LocalDateTimeModelResolve implements DataModelResolve<LocalDateTime> {
    @Override
    public boolean ableResolve(Object dataModel) {
        return dataModel instanceof LocalDateTime;
    }

    @Override
    public String resolve(LocalDateTime dataModel) {
        return dataModel.toString();
    }
}
