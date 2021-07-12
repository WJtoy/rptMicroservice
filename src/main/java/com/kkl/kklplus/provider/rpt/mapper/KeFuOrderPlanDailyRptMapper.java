package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderPlanDailySearch;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface KeFuOrderPlanDailyRptMapper {

    List<RPTKeFuOrderPlanDailyEntity> getKeFuOrderDailyList(RPTKeFuOrderPlanDailySearch search);

    /**
     * 获取客服当前的停滞订单数量
     */
    List<RPTKeFuOrderPlanDailyEntity> getKeFuPendingOrderQty(@Param("endDate") Date endDate,
                                                             @Param("keFuId") Long keFuId,
                                                             @Param("productCategoryIds") List<Long> productCategoryIds);



    List<RPTKeFuOrderPlanDailyEntity> getKeFuNoGradedOrderQty(RPTKeFuOrderPlanDailySearch search);


    Integer hasReportData(@Param("systemId") Integer systemId,
                          @Param("startDate") Long startDate,
                          @Param("endDate") Long endDate,
                          @Param("keFuIds") List<Long> keFuIds,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("quarter") String quarter);

    Integer keFuPendingOrderQty(@Param("endDate") Date endDate,
                                @Param("keFuIds") List<Long> keFuIds,
                                @Param("productCategoryIds") List<Long> productCategoryIds);

    Integer keFuNoGradedOrderQty(@Param("keFuIds") List<Long> keFuIds,
                                 @Param("productCategoryIds") List<Long> productCategoryIds);
}
