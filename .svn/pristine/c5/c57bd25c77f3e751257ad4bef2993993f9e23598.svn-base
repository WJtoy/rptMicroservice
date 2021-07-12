package com.kkl.kklplus.provider.rpt.entity;

public enum CloseOrderEfficiencyFlagEnum {
    UNFULFILLEDORDER(0,"未完成"),
    COMPLETE24HOUR(10,"24小时完成"),
    COMPLETE48HOUR(20,"48小时完成"),
    COMPLETE72HOUR(30,"72小时完成"),
    OVERCOMPLETE72HOUR(40,"超过72小时完成"),
    CANCEL24HOUR(50,"24小时内取消"),
    CANCEL48HOUR(60,"48小时内取消"),
    CANCEL72HOUR(70,"72小时内取消"),
    OVERCANCEL72HOUR(80,"超过72小时取消");

    private int value;
    private String label;

    private CloseOrderEfficiencyFlagEnum(int value, String label) {
        this.value = value;
        this.label = label;
    }

    public int getValue() {
        return this.value;
    }

    public String getLabel() {
        return this.label;
    }
}
