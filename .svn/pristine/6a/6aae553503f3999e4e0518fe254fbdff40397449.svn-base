package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTAreaOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTAreaOrderPlanDailySearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface AreaOrderPlanDailyRptMapper {

    //查询省每日下单订单
    List<RPTAreaOrderPlanDailyEntity> getProvinceOrderPlanDay(RPTAreaOrderPlanDailySearch search);

    //查询市每日下单订单
    List<RPTAreaOrderPlanDailyEntity> getCityOrderPlanDay(RPTAreaOrderPlanDailySearch search);

    //查询区每日下单订单
    List<RPTAreaOrderPlanDailyEntity> getAreaOrderPlanDay(RPTAreaOrderPlanDailySearch search);


    Integer hasReportData(RPTAreaOrderPlanDailySearch condition);
}
