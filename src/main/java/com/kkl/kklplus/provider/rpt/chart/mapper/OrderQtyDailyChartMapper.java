package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderQtyDailyChartEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

@Mapper
public interface OrderQtyDailyChartMapper {

    RPTOrderQtyDailyChartEntity getOrderQtyDailyChartData(@Param("startDate") Long startDate,
                                                          @Param("endDate") Long endDate,
                                                          @Param("systemId") Integer systemId);

    Integer getPlanOrderData(@Param("startDate") Long startDate,
                             @Param("endDate") Long endDate,
                             @Param("systemId") Integer systemId,
                             @Param("quarter") String quarter);

    Integer getCancelledOrderData(@Param("startDate") Long startDate,
                                  @Param("endDate") Long endDate,
                                  @Param("systemId") Integer systemId,
                                  @Param("quarter") String quarter);

    Integer getCompletedOrderData(@Param("startDate") Long startDate,
                                  @Param("endDate") Long endDate,
                                  @Param("systemId") Integer systemId,
                                  @Param("quarter") String quarter);

    Integer getFinancialAuditData(@Param("startDate") Long startDate,
                                  @Param("endDate") Long endDate,
                                  @Param("systemId") Integer systemId,
                                  @Param("quarter") String quarter);

    Integer getAutoFinancialAuditData(@Param("startDate") Long startDate,
                                      @Param("endDate") Long endDate,
                                      @Param("systemId") Integer systemId);

    Integer getAbnormalOrderData(@Param("startDate") Long startDate,
                                 @Param("endDate") Long endDate,
                                 @Param("systemId") Integer systemId);

    Integer getAutoCompletedData(@Param("startDate") Long startDate,
                                 @Param("endDate") Long endDate,
                                 @Param("systemId") Integer systemId,
                                 @Param("quarter") String quarter);

    Integer getUnCompletedOrderData(@Param("startDate") Long startDate,
                                    @Param("endDate") Long endDate,
                                    @Param("systemId") Integer systemId);

    void insertOrderQtyDailyData(RPTOrderQtyDailyChartEntity entity);

    void deleteOrderQtyDailyFromRptDB(@Param("systemId") Integer systemId,
                                        @Param("startDate") Long startDate,
                                        @Param("endDate") Long endDate);
}
