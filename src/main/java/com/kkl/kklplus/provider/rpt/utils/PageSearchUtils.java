package com.kkl.kklplus.provider.rpt.utils;

import com.google.common.collect.Lists;

import java.util.List;

/**
 * @author: Zhoucy
 * @date: 2020/9/11
 * @Description:
 */
public class PageSearchUtils {

    public static <T> List<T> exec(QueryAction<T> action, int pageSize) {
        List<T> result = Lists.newArrayList();
        int pageNo = 1;
        int startLimit = 0;
        int size = 0;
        do {
            startLimit = (pageNo - 1) * pageSize;
            List<T> list = action.query(startLimit, pageSize);
            if (list != null) {
                result.addAll(list);
                size = list.size();
            }
            pageNo++;
        } while (size == pageSize);

        return result;
    }

    public interface QueryAction<T> {
        <T> List<T> query(int startLimit, int pageSize);
    }
}
