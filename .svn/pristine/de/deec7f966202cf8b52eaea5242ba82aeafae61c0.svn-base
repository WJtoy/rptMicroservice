package com.kkl.kklplus.provider.rpt.entity;

import lombok.Getter;
import lombok.Setter;

public class ChargeDailyEntity {

    @Getter
    @Setter
    private Long countNum;

    @Getter
    @Setter
    private int dayIndex;

    @Getter
    @Setter
    private int yearMonth;


    public String getYearMonthDay() {
        if(dayIndex <= 9){
            return String.valueOf(yearMonth) + "0" + String.valueOf(dayIndex);
        }else {
            return String.valueOf(yearMonth) + String.valueOf(dayIndex);
        }
    }
}
