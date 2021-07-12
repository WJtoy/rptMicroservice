package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

public class CustomerPerformanceRptEntity extends RPTBase {

    @Getter
    @Setter
    private Integer yearMonth;

    @Getter
    @Setter
    private Long productCategoryId;

    @Getter
    @Setter
    private Long customerId = 0L;

    @Getter
    @Setter
    private Integer createQty = 0;

    @Getter
    @Setter
    private Integer finishQty = 0;

    @Getter
    @Setter
    private Integer noFinishQty = 0;

    @Getter
    @Setter
    private Integer returnQty = 0;


    @Getter
    @Setter
    private Integer cancelQty = 0;

}
