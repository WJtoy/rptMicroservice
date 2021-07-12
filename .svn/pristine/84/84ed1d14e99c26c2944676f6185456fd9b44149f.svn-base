package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface CustomerOrderPlanDailyRptMapper {

    List<RPTCustomerOrderPlanDailyEntity> getCustomerOrderPlanDailyList(RPTCustomerOrderPlanDailySearch search);

    Long getCustomerOrderPlanMonth(RPTCustomerOrderPlanDailySearch search);

    Integer hasReportData(RPTCustomerOrderPlanDailySearch search);
}
