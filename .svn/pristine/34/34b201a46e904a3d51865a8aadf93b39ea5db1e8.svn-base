package com.kkl.kklplus.provider.rpt.chart.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

public class RPTServicePointQtyEntity extends RPTBase {

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
     * 常用网点数量
     */
    @Getter
    @Setter
    private Long frequentQty = 0L;

    /**
     * 试用网点数量
     */
    @Getter
    @Setter
    private Long trialQty = 0L;


    /**
     * 自动派单的网点数量
     */
    @Getter
    @Setter
    private Long autoPlanQty = 0L;

    /**
     * 创建时间
     */
    @Getter
    @Setter
    private Long createDate;

    /**
     * 网点数量比率
     */
    @Getter
    @Setter
    private String totalRate = "0";

    public Long getTotal() {
        return frequentQty + trialQty;
    }
}
