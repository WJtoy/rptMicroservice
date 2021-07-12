package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface DepositRechargeSummaryMapper {

    List<RPTCustomerRechargeSummaryEntity> getOnlineRechargeSummaryList(@Param("customerId") Long customerId,
                                                                        @Param("beginDate") Date beginDate,
                                                                        @Param("endDate") Date endDate,
                                                                        @Param("quarter") String quarter);
    //获取订单完成扣款
    List<RPTCustomerRechargeSummaryEntity> getOrderCompleteDeductSummaryList(@Param("customerId") Long customerId,
                                                                        @Param("beginDate") Date beginDate,
                                                                        @Param("endDate") Date endDate,
                                                                        @Param("quarter") String quarter);


    Page<RPTRechargeRecordEntity> getDepositRechargeDetailsPage(@Param("customerId") Long customerId,
                                                                @Param("beginDate") Date beginDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("actionType") Integer actionType,
                                                                @Param("quarters") List<String> quarters,
                                                                @Param("pageNo") Integer pageNo,
                                                                @Param("pageSize") Integer pageSize);


    List<RPTRechargeRecordEntity> getDepositRechargeDetailsList(@Param("customerId") Long customerId,
                                                                @Param("beginDate") Date beginDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("actionType") Integer actionType,
                                                                @Param("quarters") List<String> quarters);


    Integer hasReportData(@Param("customerId") Long customerId,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("quarter") String quarter);


    Long hasDepositReportData(@Param("customerId") Long customerId,
                              @Param("beginDate") Date beginDate,
                              @Param("endDate") Date endDate,
                              @Param("actionType") Integer actionType,
                              @Param("quarters") List<String> quarters);


}
