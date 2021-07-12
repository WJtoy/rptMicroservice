package com.kkl.kklplus.provider.rpt.ms.sys.enums;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.web.RPTDict;

import java.util.List;

/**
 * 结算方式
 */
public enum PaymentType {

    /**
     * 结算方式
     */
    MONTHLY(10, "月结"),
    IMMEDIATELY(20, "即结"),
    BEFOREHAND(30, "预付");

    public int value;
    public String label;

    PaymentType(int value, String label) {
        this.value = value;
        this.label = label;
    }

    public static List<RPTDict> getAllPaymentTypes() {
        List<RPTDict> list = Lists.newArrayList();
        list.add(new RPTDict(MONTHLY.value, MONTHLY.label));
        list.add(new RPTDict(IMMEDIATELY.value, IMMEDIATELY.label));
        list.add(new RPTDict(BEFOREHAND.value, BEFOREHAND.label));
        return list;
    }

}
