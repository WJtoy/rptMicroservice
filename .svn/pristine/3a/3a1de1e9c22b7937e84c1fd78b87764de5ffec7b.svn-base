package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuOrderCancelledDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderCancelledDailySearch;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface KeFuOrderCancelledDailyRptMapper {

    List<RPTKeFuOrderCancelledDailyEntity> getKeFuCancelledOrderList(RPTKeFuOrderCancelledDailySearch search);

    Long getKeFuOrderCancelledMonth(@Param("systemId") Integer systemId,
                                    @Param("startDate") Long startDate,
                                    @Param("endDate") Long endDate,
                                    @Param("keFuIds") List<Long> keFuIds,
                                    @Param("productCategoryIds") List<Long> productCategoryIds,
                                    @Param("quarter") String quarter);

    Integer hasReportData(@Param("systemId") Integer systemId,
                          @Param("startDate") Long startDate,
                          @Param("endDate") Long endDate,
                          @Param("keFuIds") List<Long> keFuIds,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("quarter") String quarter);

}
