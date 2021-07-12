package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RptCustomerMonthOrderEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface CustomerMonthPlanDailyRptMapper {

    List<RptCustomerMonthOrderEntity> getCustomerMonthPlanDailyList(@Param("startYearMonth") Integer startYearMonth,
                                                                    @Param("endYearMonth") Integer endYearMonth,
                                                                    @Param("customerId")Long customerId,
                                                                    @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                    @Param("salesId") Long salesId,
                                                                    @Param("systemId") Integer systemId,
                                                                    @Param("subFlag") Integer subFlag);

    Integer hasReportData(@Param("startYearMonth") Integer startYearMonth,
                          @Param("endYearMonth") Integer endYearMonth,
                          @Param("customerId")Long customerId,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("salesId") Long salesId,
                          @Param("systemId") Integer systemId,
                          @Param("subFlag") Integer subFlag);
}
