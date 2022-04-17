package com.lj.wordmark.modelresolve;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public interface DataModelResolve<T> {

    boolean ableResolve(Object dataModel);

    String resolve(T dataModel);


}
