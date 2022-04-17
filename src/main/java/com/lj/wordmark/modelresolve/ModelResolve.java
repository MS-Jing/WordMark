package com.lj.wordmark.modelresolve;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
public class ModelResolve {
    private static final List<DataModelResolve> RESOLVES = new ArrayList<>();

    static {
        RESOLVES.add(new NullDataModel());
        RESOLVES.add(new StringDataModel());
        RESOLVES.add(new BaseTypeModel());
        RESOLVES.add(new LocalDateModelResolve());
        RESOLVES.add(new LocalDateTimeModelResolve());
        RESOLVES.add(new DateModelResolve());
    }

    public static boolean ableResolve(Object dataModel) {
        for (DataModelResolve resolve : RESOLVES) {
            if (resolve.ableResolve(dataModel)) {
                return true;
            }
        }
        return false;
    }

    public static String resolve(Object dataModel) {
        for (DataModelResolve resolve : RESOLVES) {
            if (resolve.ableResolve(dataModel)) {
                return resolve.resolve(dataModel);
            }
        }
        return null;
    }

    public static void addResolve(DataModelResolve resolve) {
        RESOLVES.add(resolve);
    }
}
