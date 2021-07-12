package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuAverageOrderFeeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTReminderResponseTimeSearch;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface KeFuAverageOrderFeeRptMapper {

    List<RPTKeFuAverageOrderFeeEntity> getSpecialOrderFeeList (@Param("systemId") Integer systemId,
                                                               @Param("yearMonth") Integer yearMonth,
                                                               @Param("quarter") String quarter,
                                                               @Param("productCategoryIds") List<Long> productCategoryIds,
                                                               @Param("customerId") Long customerId,
                                                               @Param("areaType") Integer areaType,
                                                               @Param("areaId") Long areaId);

    List<RPTKeFuAverageOrderFeeEntity> getGradedOrderData (@Param("systemId") Integer systemId,
                                                           @Param("yearMonth") Integer yearMonth,
                                                           @Param("quarter") String quarter,
                                                           @Param("productCategoryIds") List<Long> productCategoryIds,
                                                           @Param("customerId") Long customerId,
                                                           @Param("areaType") Integer areaType,
                                                           @Param("areaId") Long areaId);


    List<Long> getKeFuCustomer (@Param("keFuId") Long keFuId,@Param("systemId") Integer systemId);

    List<Long> getVipCustomer (@Param("systemId") Integer systemId);

    //市客服
    List<RPTKeFuAverageOrderFeeEntity> getKeFuAreaA(@Param("systemId") Integer systemId,
                                                    @Param("keFuId") Long keFuId);

    //省客服
    List<RPTKeFuAverageOrderFeeEntity> getKeFuAreaB(@Param("systemId") Integer systemId,
                                                    @Param("keFuId") Long keFuId);

    //区客服
    List<RPTKeFuAverageOrderFeeEntity> getKeFuAreaD(@Param("systemId") Integer systemId,
                                                    @Param("keFuId") Long keFuId);

    //全国客服
    Integer getKeFuAreaC(@Param("systemId") Integer systemId,
                         @Param("keFuId") Long keFuId);

    List<RPTKeFuAverageOrderFeeEntity> getOrderSum(@Param("systemId") Integer systemId,
                                                      @Param("yearMonth") Integer yearMonth,
                                                      @Param("quarter") String quarter,
                                                      @Param("productCategoryIds") List<Long> productCategoryIds,
                                                      @Param("customerId") Long customerId,
                                                      @Param("areaType") Integer areaType,
                                                      @Param("areaId") Long areaId);

    List<RPTKeFuAverageOrderFeeEntity> getVipOrderSum(@Param("systemId") Integer systemId,
                                                      @Param("yearMonth") Integer yearMonth,
                                                      @Param("quarter") String quarter,
                                                      @Param("productCategoryIds") List<Long> productCategoryIds,
                                                      @Param("customerId") Long customerId,
                                                      @Param("areaType") Integer areaType,
                                                      @Param("areaId") Long areaId,
                                                      @Param("keFuId") Long keFuId);
}
