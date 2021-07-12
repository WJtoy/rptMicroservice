package com.kkl.kklplus.provider.rpt.utils;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class NumberUtils {

    public static int toInteger(Number value) {
        return value == null ? 0 : value.intValue();
    }

    public static long toLong(Number value) {
        return value == null ? 0 : value.longValue();
    }

    public static double toDouble(Number value) {
        return value == null ? 0 : value.doubleValue();
    }
}
