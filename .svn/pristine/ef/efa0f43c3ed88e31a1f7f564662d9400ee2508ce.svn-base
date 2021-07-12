package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CustomerRechargeSummaryMapper {

    List<RPTCustomerRechargeSummaryEntity> getOnlineRechargeSummaryList(@Param("customerId") Long customerId,
                                                                          @Param("beginDate") Date beginDate,
                                                                          @Param("endDate") Date endDate,
                                                                          @Param("quarter") String quarter);

    List<RPTCustomerRechargeSummaryEntity> getFinancialRechargeSummaryList(@Param("customerId") Long customerId,
                                                                           @Param("beginDate") Date beginDate,
                                                                           @Param("endDate") Date endDate,
                                                                           @Param("quarter") String quarter);


    List<RPTCustomerRechargeSummaryEntity> getRechargeSummaryList(@Param("customerId") Long customerId,
                                                                           @Param("beginDate") Date beginDate,
                                                                           @Param("endDate") Date endDate,
                                                                           @Param("quarter") String quarter);

    Integer hasReportData(@Param("customerId") Long customerId,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("quarter") String quarter);


}
