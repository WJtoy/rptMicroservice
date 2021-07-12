/**
 * Copyright &copy; 2012-2016 <a href="https://github.com/thinkgem/jeesite">JeeSite</a> All rights reserved.
 */
package com.kkl.kklplus.provider.rpt.utils;

import java.util.List;
import java.util.UUID;
import java.util.stream.Collectors;

public class StringUtils extends org.apache.commons.lang3.StringUtils {

    public static String toString(String str) {
        return isNotEmpty(str) ? str : "";
    }

    /**
     * 转换为Double类型
     */
    public static Double toDouble(Object val) {
        if (val == null) {
            return 0D;
        }
        try {
            return Double.valueOf(trim(val.toString()));
        } catch (Exception e) {
            return 0D;
        }
    }

    /**
     * 转换为Float类型
     */
    public static Float toFloat(Object val) {
        return toDouble(val).floatValue();
    }

    /**
     * 原来的toLong()方法出现过精度不够的问题，故新增该方法
     */
    public static Long toLong(Object val) {
        if (val == null) {
            return 0L;
        }
        if (isBlank(val.toString())) {
            return 0L;
        }
        try {
            return Long.valueOf(trim(val.toString()));
        } catch (Exception e) {
            return 0L;
        }
    }

    /**
     * 转换为Integer类型
     */
    public static Integer toInteger(Object val) {
        return toLong(val).intValue();
    }


    /**
     * 获取UUID
     *
     * @return
     */
    public static String uuid() {
        return UUID.randomUUID().toString().replaceAll("-", "");
    }

    /**
     * 将List转换成字符串序列
     *
     * @param strings 字符串列表
     */
    public static String joinStringList(List<String> strings) {
        String result = "";
        if (strings != null && !strings.isEmpty()) {
            result = strings.stream().filter(StringUtils::isNotBlank).collect(Collectors.joining(","));
        }
        return result;
    }

}
