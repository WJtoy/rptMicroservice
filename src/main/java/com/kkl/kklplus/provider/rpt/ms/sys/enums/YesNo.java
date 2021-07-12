package com.kkl.kklplus.provider.rpt.ms.sys.enums;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.web.RPTDict;

import java.util.List;

/**
 * 是否
 */
public enum YesNo {
    /**
     * 是否标记
     */
    YES(1, "是"),
    NO(0, "否");

    public int value;
    public String label;

    YesNo(int value, String label) {
        this.value = value;
        this.label = label;
    }

    public static List<RPTDict> getAllYesNo() {
        List<RPTDict> list = Lists.newArrayList();
        list.add(new RPTDict(YES.value, YES.label));
        list.add(new RPTDict(NO.value, NO.label));
        return list;
    }
}
