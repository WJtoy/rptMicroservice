package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.abnormal.AbnormalFinancialAuditStatistics;
import com.kkl.kklplus.entity.rpt.RPTAbnormalFinancialAuditEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.entity.AbnormalFinancialReviewEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;
import java.util.Map;

@Mapper
public interface AbnormalFinancialReviewRptMapper {

    List<AbnormalFinancialReviewEntity> getManualChargeList(@Param("startDate") Date startDate,
                                                            @Param("endDate") Date endDate);

    /**
     * 获取自动对账的数量
     */
    List<AbnormalFinancialReviewEntity> getAutoChargeList(@Param("startDate") Date startDate,
                                                            @Param("endDate") Date endDate);


    List<AbnormalFinancialAuditStatistics> financialAuditStat(@Param("createBeginDt") Long createBeginDt,
                                                              @Param("createEndDt")Long createEndDt);

    void insertAbnormalFinancialDaily(AbnormalFinancialReviewEntity entity);

    void deleteAbnormalFinancialData(@Param("startDate") Long startDate,
                                     @Param("endDate") Long endDate,
                                     @Param("systemId") Integer systemId);


    List<RPTAbnormalFinancialAuditEntity> getManualMonthList(RPTCustomerOrderPlanDailySearch search);

    List<RPTAbnormalFinancialAuditEntity> getCheckerManualList(@Param("startDate") Long startDate,
                                                               @Param("endDate") Long endDate,
                                                               @Param("systemId") Integer systemId,
                                                               @Param("abnormalIds") List<Long> abnormalIds);

    List<RPTAbnormalFinancialAuditEntity> getAutoMonthList(RPTCustomerOrderPlanDailySearch search);

    List<RPTAbnormalFinancialAuditEntity> getAbnormalList(RPTCustomerOrderPlanDailySearch search);

    List<RPTAbnormalFinancialAuditEntity> getCheckerAbnormalList(RPTCustomerOrderPlanDailySearch search);

    List<RPTAbnormalFinancialAuditEntity> getCompletedDaily(@Param("startYearMonth") Integer startYearMonth,
                                                            @Param("systemId") Integer systemId);

    Integer hasReportData(@Param("startYearMonth") Integer startYearMonth,
                          @Param("systemId") Integer systemId);

    Integer hasManualReportData(RPTCustomerOrderPlanDailySearch search);




}
