package com.kdgc.worddemo.util;

import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * @Description: CollectionUtil工具类
 * @Author: LeMenPan
 * @Date:   2021-01-17
 * @Version: V-1.0
 */
public class CollectionUtilsPan {

    /**
     * 使用属性过滤集合对象重复数据
     * @param keyExtractor
     * @param <T>
     * @return
     */
    public static <T> Predicate<T> distinctByKey(Function<? super T, Object> keyExtractor) {
        Map<Object, Boolean> seen = new ConcurrentHashMap<>();
        return t -> seen.putIfAbsent(keyExtractor.apply(t), Boolean.TRUE) == null;
    }
}


