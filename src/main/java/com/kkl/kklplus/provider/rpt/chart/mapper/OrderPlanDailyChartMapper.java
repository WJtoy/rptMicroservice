package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderPlanDailyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface OrderPlanDailyChartMapper {


    List<RPTOrderPlanDailyEntity> getOrderPlanDailyData(@Param("systemId") Integer systemId,
                                                        @Param("startDate") Long startDate,
                                                        @Param("endDate") Long endDate,
                                                        @Param("quarters") List<String> quarters);

    List<RPTOrderPlanDailyEntity> getProductCategoryPlanDailyData(@Param("systemId") Integer systemId,
                                                                  @Param("startDate") Long startDate,
                                                                  @Param("endDate") Long endDate,
                                                                  @Param("quarters") List<String> quarters);

    List<RPTOrderPlanDailyEntity> getCustomerPlanQtyData(@Param("systemId") Integer systemId,
                                                         @Param("startDate") Long startDate,
                                                         @Param("endDate") Long endDate,
                                                         @Param("quarters") List<String> quarters);

    List<RPTOrderPlanDailyEntity> getProductCategoryCustomerPlanQty(@Param("systemId") Integer systemId,
                                                                    @Param("startDate") Long startDate,
                                                                    @Param("endDate") Long endDate,
                                                                    @Param("quarters") List<String> quarters);


}
