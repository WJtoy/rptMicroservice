package com.kkl.kklplus.provider.rpt.chart.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

public class RPTServicePointStreetQtyEntity extends RPTBase{

    /**
     * 服务品类
     */
    @Getter
    @Setter
    private Long productCategoryId = 0L;


    /**
     * 服务品类
     */
    @Getter
    @Setter
    private String productCategoryName = "";

    /**
     * 网点街道总数量
     */
    @Getter
    @Setter
    private Long servicePointStreet = 0L;

    /**
     * 常用网点街道数量
     */
    @Getter
    @Setter
    private Long frequentServicePoint = 0L;
    /**
     * 常用网点街道比率
     */
    @Getter
    @Setter
    private String frequentServicePointRate = "0";

    /**
     * 试用网点街道数量
     */
    @Getter
    @Setter
    private Long trialServicePoint = 0L;

    /**
     * 试用网点街道比率
     */
    @Getter
    @Setter
    private String trialServicePointRate = "0";
    /**
     * 无网点街道数量
     */
    @Getter
    @Setter
    private Long withoutServicePoint = 0L;

    /**
     * 无网点街道比率
     */
    @Getter
    @Setter
    private String withoutServicePointRate = "0";

    /**
     * 自动派单的街道数量
     */
    @Getter
    @Setter
    private Long autoPlanStreet = 0L;

    /**
     * 自动派单的街道比率
     */
    @Getter
    @Setter
    private String autoPlanStreetRate = "0";

    /**
     * 自动派单的街道无常用网点街道的数量
     */
    @Getter
    @Setter
    private Long autoPlanWithoutFrequent = 0L;

    /**
     * 常用网点街道无自动派单
     */
    @Getter
    @Setter
    private Long frequentWithoutAutoPlan = 0L;

    /**
     * 创建时间
     */
    @Getter
    @Setter
    private Long createDate;



}
