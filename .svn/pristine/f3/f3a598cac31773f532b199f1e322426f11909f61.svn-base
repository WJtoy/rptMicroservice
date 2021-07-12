package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderCrushQtyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface OrderCrushQtyChartMapper {

    RPTOrderCrushQtyEntity getOrderCrushQtyData(@Param("systemId") Integer systemId,
                                                @Param("startDate") Long startDate,
                                                @Param("endDate") Long endDate);
    Integer getPlainOrderQty(@Param("systemId") Integer systemId,
                             @Param("startDate") Long startDate,
                             @Param("endDate") Long endDate);

    Integer getOrderCrushQty(@Param("startDate") Date startDate,
                             @Param("endDate") Date endDate);

    Integer getOnceCrushQty(@Param("startDate") Date startDate,
                            @Param("endDate") Date endDate);

    Integer getRepeatedlyCrushQty(@Param("startDate") Date startDate,
                                  @Param("endDate") Date endDate);


    Integer getCompletedCrushQty(@Param("startDate") Date startDate,
                                 @Param("endDate") Date endDate);

    Integer getCompletedOnceCrush(@Param("startDate") Date startDate,
                                  @Param("endDate") Date endDate);

    Integer getCompletedRepeatedlyCrush(@Param("startDate") Date startDate,
                                        @Param("endDate") Date endDate);


    List<RPTOrderCrushQtyEntity> getCompletedOrderCrush(@Param("startDate") Date startDate,
                                                        @Param("endDate") Date endDate);

    void insertOrderCrushQtyData(RPTOrderCrushQtyEntity entity);

    void deleteOrderCrushQtyFromRptDB(@Param("systemId") Integer systemId,
                                      @Param("startDate") Long startDate,
                                      @Param("endDate") Long endDate);
}
