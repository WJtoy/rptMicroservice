package com.kkl.kklplus.provider.rpt.entity;


import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

public class ComplainDailyEntity  extends RPTBase  {

    /**
     * 投诉ID
     */
    @Getter
    @Setter
    private Long complainId;


    /**
     * 客户ID
     */
    @Getter
    @Setter
    private Long customerId;



    /**
     * 品类ID
     */
    @Getter
    @Setter
    private Long productCategoryId;


    /**
     * 省ID
     */
    @Getter
    @Setter
    private Long provinceId;


    /**
     * 市ID
     */
    @Getter
    @Setter
    private Long cityId;

    /**
     * 区ID
     */
    @Getter
    @Setter
    private Long countyId;


    /**
     * 投诉时间
     */
    @Getter
    @Setter
    private Long complainDt;


    /**
     * 投诉时间
     */
    @Getter
    @Setter
    private Date complainDate;


    /**
     * 天数
     */
    @Getter
    @Setter
    private Integer dayIndex;


    /**
     * 判断对象
     */
    @Getter
    @Setter
    private Integer judgeObject;


    /**
     * 责任项目
     */
    @Getter
    @Setter
    private Integer judgeItem;


    /**
     * 状态
     */
    @Getter
    @Setter
    private Integer status;


    /**
     * 投诉状态
     */
    @Getter
    @Setter
    private Integer complainStatus = 0;


    /**
     * 中差评判断
     */
    @Getter
    @Setter
    private Integer evaluateStatus = 0;



}
