package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface ServicePointBalanceRptMapper {

    Page<RPTServicePointBalanceEntity> getServicePointPayableData(@Param("servicePointId") Long servicePointId,
                                                                  @Param("paymentType") Integer paymentType,
                                                                  @Param("selectedYear") Integer selectedYear,
                                                                  @Param("selectedMonth") Integer selectedMonth,
                                                                  @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                  @Param("pageNo")  Integer pageNo,
                                                                  @Param("pageSize")  Integer pageSize);


    List<RPTServicePointBalanceEntity> getServicePointPayableReport(@Param("servicePointId") Long servicePointId,
                                                                  @Param("paymentType") Integer paymentType,
                                                                  @Param("selectedYear") Integer selectedYear,
                                                                  @Param("selectedMonth") Integer selectedMonth,
                                                                  @Param("productCategoryIds") List<Long> productCategoryIds);
    /**
     * 查询网点指定年月的余额
     */
    List<RPTServicePointBalanceEntity> getServicePointBalanceData(@Param("servicePointIds") List<Long> servicePointIds,
                                                                  @Param("paymentType") Integer paymentType,
                                                                  @Param("selectedYear") Integer selectedYear,
                                                                  @Param("selectedMonth") Integer selectedMonth,
                                                                  @Param("productCategoryIds") List<Long> productCategoryIds);


    Integer  getServicePointPayableDataSum(@Param("servicePointId") Long servicePointId,
                                           @Param("paymentType") Integer paymentType,
                                           @Param("selectedYear") Integer selectedYear,
                                           @Param("selectedMonth") Integer selectedMonth,
                                           @Param("productCategoryIds") List<Long> productCategoryIds);

    Integer  getServicePointBalanceDataSum(@Param("servicePointIds") List<Long> servicePointIds,
                                           @Param("paymentType") Integer paymentType,
                                           @Param("selectedYear") Integer selectedYear,
                                           @Param("selectedMonth") Integer selectedMonth,
                                           @Param("productCategoryIds") List<Long> productCategoryIds);


}
