package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

public class AbnormalFinancialReviewEntity extends RPTBase {

    /**
     * 创建人ID
     */
    @Getter
    @Setter
    private Long createId;


    /**
     * 创建人
     */
    @Getter
    @Setter
    private String createName;

    /**
     * 异常单类型
     */
    @Getter
    @Setter
    private Integer subType;


    /**
     * 统计类型
     */
    @Getter
    @Setter
    private Integer type;


    /**
     * 审单时间
     */
    @Getter
    @Setter
    private Long auditTime;

    /**
     * 品类
     */
    @Getter
    @Setter
    private Long productCategoryId;


    /**
     * 下单数量
     */
    @Getter
    @Setter
    private Integer qty;


    /**
     * 天
     */
    @Getter
    @Setter
    private Integer dayIndex;




}
