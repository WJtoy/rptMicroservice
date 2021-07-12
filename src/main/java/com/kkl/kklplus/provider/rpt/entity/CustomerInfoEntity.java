package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

public class CustomerInfoEntity extends RPTBase {


    /**
     * 客户ID
     */
    @Getter
    @Setter
    private Long customerId;



    /**
     *  vip标识
     */
    @Getter
    @Setter
    private Integer vipFlag;
}
