package com.kkl.kklplus.provider.rpt.entity;

import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.web.RPTDict;

import java.util.Map;

public enum  ServicePointStatusEnum {

    NORMAL(10, "正常"),
    PAUSED(20, "暂停派单"),
    CANCELLED(30, "取消合作"),
    BLACKLIST(100, "黑名单");

    private int value;
    private String label;

    ServicePointStatusEnum(int value, String label) {
        this.value = value;
        this.label = label;
    }

    public int getValue() {
        return value;
    }

    public String getLabel() {
        return label;
    }

    private static final Map<Integer, ServicePointStatusEnum> map = Maps.newHashMap();

    static {
        for (ServicePointStatusEnum item : ServicePointStatusEnum.values()) {
            map.put(item.value, item);
        }
    }

    public static RPTDict createDict(ServicePointStatusEnum status) {
        if (status == null) {
            return new RPTDict(0, "");
        } else {
            return new RPTDict(status.getValue(), status.getLabel());
        }
    }

    public static ServicePointStatusEnum valueOf(Integer value) {
        ServicePointStatusEnum status = null;
        if (value != null) {
            status = map.get(value);
        }
        return status;
    }
}
