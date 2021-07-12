package com.kkl.kklplus.provider.rpt.ms.sys.enums;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.web.RPTDict;

import java.util.List;

public enum OrderStatusType {

    /**
     * 工单状态
     */
    NEW(10, "下单"),
    APPROVED(20, "待接单"),
    ACCEPTED(30, "已接单"),
    PLANNED(40, "已派单"),
    SERVICED(50, "已上门"),
    APP_COMPLETED(55, "待回访"),
    RETURNING(60, "退单申请"),
    CANCELING(70, "取消中"),
    COMPLETED(80, "完成"),
    CHARGED(85, "已入账"),
    RETURNED(90, "已退单"),
    CANCELED(100, "已取消");

    public int value;
    public String label;

    OrderStatusType(int value, String label) {
        this.value = value;
        this.label = label;
    }

    public static List<RPTDict> getAllOrderStatusTypes() {
        List<RPTDict> list = Lists.newArrayList();
        list.add(new RPTDict(NEW.value, NEW.label));
        list.add(new RPTDict(APPROVED.value, APPROVED.label));
        list.add(new RPTDict(ACCEPTED.value, ACCEPTED.label));
        list.add(new RPTDict(PLANNED.value, PLANNED.label));
        list.add(new RPTDict(SERVICED.value, SERVICED.label));
        list.add(new RPTDict(APP_COMPLETED.value, APP_COMPLETED.label));
        list.add(new RPTDict(RETURNING.value, RETURNING.label));
        list.add(new RPTDict(CANCELING.value, CANCELING.label));
        list.add(new RPTDict(COMPLETED.value, COMPLETED.label));
        list.add(new RPTDict(CHARGED.value, CHARGED.label));
        list.add(new RPTDict(RETURNED.value, RETURNED.label));
        list.add(new RPTDict(CANCELED.value, CANCELED.label));
        return list;
    }
}
