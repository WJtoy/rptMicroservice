package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerRevenueEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CustomerRevenueRptMapper {

    List<RPTCustomerRevenueEntity> getCustomerRevenueList(@Param("systemId") Integer systemId,
                                                          @Param("customerId") Long customerId,
                                                          @Param("startDate") Long startDate,
                                                          @Param("endDate") Long endDate,
                                                          @Param("productCategoryIds") List<Long> productCategoryIds,
                                                          @Param("quarter") String quarter);

    List<RPTCustomerRevenueEntity> getFinishOrderData(@Param("startDate") Date startDate,
                                                      @Param("endDate") Date endDate,
                                                      @Param("quarter") String quarter);

    List<RPTCustomerRevenueEntity> getReceivableCharge(@Param("startDate") Date startDate,
                                                       @Param("endDate") Date endDate,
                                                       @Param("quarter") String quarter);

    List<RPTCustomerRevenueEntity> getPayableChargeA(@Param("startDate") Date startDate,
                                                       @Param("endDate") Date endDate,
                                                       @Param("quarter") String quarter);

    List<RPTCustomerRevenueEntity> getPayableChargeB(@Param("startDate") Date startDate,
                                                     @Param("endDate") Date endDate,
                                                     @Param("quarter") String quarter);

    void insertCustomerRevenueData(@Param("list") List<RPTCustomerRevenueEntity> entity);


    void deleteCustomerRevenueFromRptDB(@Param("systemId") Integer systemId,
                                        @Param("startDate") Long startDate,
                                        @Param("endDate") Long endDate,
                                        @Param("quarter") String quarter);

    Integer hasReportData(@Param("systemId") Integer systemId,
                          @Param("startDate") Long startDate,
                          @Param("endDate") Long endDate,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("quarter") String quarter);
}
