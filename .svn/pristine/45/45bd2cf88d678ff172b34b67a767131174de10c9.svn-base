package com.kkl.kklplus.provider.rpt.chart.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

@Mapper
public interface CustomerComplainChartMapper {

    Integer getCustomerComplainQty(@Param("systemId") Integer systemId,
                                   @Param("beginDate") Long beginDate,
                                   @Param("endDate") Long endDate,
                                   @Param("quarter") String quarter);

    Integer getCustomerValidComplain(@Param("systemId") Integer systemId,
                                     @Param("beginDate") Long beginDate,
                                     @Param("endDate") Long endDate,
                                     @Param("quarter") String quarter);

    Integer getCustomerMediumPoorEvaluate(@Param("systemId") Integer systemId,
                                          @Param("beginDate") Long beginDate,
                                          @Param("endDate") Long endDate,
                                          @Param("quarter") String quarter);
}
