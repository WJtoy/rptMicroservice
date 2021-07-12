package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTFinancialReviewDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface FinancialReviewDetailsRptMapper {

    Page<RPTFinancialReviewDetailsEntity> getFinancialReviewDetailsList(RPTCustomerOrderPlanDailySearch search);

    List<RPTFinancialReviewDetailsEntity> getFinancialReviewDetailsByList(RPTCustomerOrderPlanDailySearch search);

    Integer hasReportData(RPTCustomerOrderPlanDailySearch search);


}
