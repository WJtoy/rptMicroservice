package com.kkl.kklplus.provider.rpt.utils;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.search.RPTSearchBase;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.function.Function;

@Slf4j
public class PageUtils {

    /**
     * 获取单笔记录数据
     */
    public static <R extends RPTBase, C extends RPTSearchBase> List<R> getReturnPage(C search, Function<C, List<R>> function) {
        List<R> result = function.apply(search);
        return result;
    }

    /**
     * 计算总页数
     */
    public static int calcPageCount(int rowCount, int pageSize) {
        return (rowCount + pageSize - 1) / pageSize;
    }

    /**
     * 计算分页语句 LIMIT的offset值
     */
    public static int calcLimitOffsetValue(int pageSize, int pageIndex) {
        return (pageIndex - 1) * pageSize;
    }
}
