package com.kulang.commonUtils;

/**
 * Created by wenqiang.wang on 2017/3/15.
 */
public interface Rule<T> {
    boolean apply(T t);
}