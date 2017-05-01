package com.kulang.commonUtils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Created by wenqiang.wang on 2017/3/15.
 */
public class ListUtils {
    public static <T> List<T> filter(Collection<T> collection, Rule<T> rule) {
        List<T> list = new ArrayList<T>();
        for (T t : collection) {
            if (rule.apply(t)) {
                list.add(t);
            }
        }

        return list;
    }

    public static <T> T get(Collection<T> collection, Rule<T> rule) {
        for (T t : collection) {
            if (rule.apply(t)) {
                return t;
            }
        }

        return null;
    }
}


