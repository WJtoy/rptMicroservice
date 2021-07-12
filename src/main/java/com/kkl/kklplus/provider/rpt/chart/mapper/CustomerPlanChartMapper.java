package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface CustomerPlanChartMapper {

    Integer getPlanOrderData(@Param("startDate") Long startDate,
                             @Param("endDate") Long endDate,
                             @Param("systemId") Integer systemId,
                             @Param("quarter") String quarter);

    List<TwoTuple> getPlanOrderProductCategoryData(@Param("startDate") Long startDate,
                                                   @Param("endDate") Long endDate,
                                                   @Param("systemId") Integer systemId,
                                                   @Param("quarter") String quarter);

}
