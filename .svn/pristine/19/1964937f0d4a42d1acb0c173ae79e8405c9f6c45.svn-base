package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface RechargeRecordRptMapper {

    Page<RPTRechargeRecordEntity> rechargeRecordPage(@Param("customerId") Long customerId,
                                                     @Param("beginDate") Date beginDate,
                                                     @Param("endDate") Date endDate,
                                                     @Param("actionType") Integer actionType,
                                                     @Param("quarters")  List<String> quarters,
                                                     @Param("pageNo")  Integer pageNo,
                                                     @Param("pageSize")  Integer pageSize);

    List<RPTRechargeRecordEntity> rechargeRecordList(@Param("customerId") Long customerId,
                                                     @Param("beginDate") Date beginDate,
                                                     @Param("endDate") Date endDate,
                                                     @Param("actionType") Integer actionType,
                                                     @Param("quarters")  List<String> quarters);

    Long hasReportData(@Param("customerId") Long customerId,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("actionType") Integer actionType,
                          @Param("quarters")  List<String> quarters);

}
