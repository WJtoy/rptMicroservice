package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CustomerReminderRptMapper {
    //获取所有客户每日下单数量

    List<RPTCustomerReminderEntity> getOrderNewQtyList(@Param("quarter") String quarter,
                                                       @Param("beginDate") Date beginDate,
                                                       @Param("endDate") Date endDate);

    //获取客户每日催单数量
    List<RPTCustomerReminderEntity> getReminderQtyList(@Param("beginDate") Long beginDate,
                                                       @Param("endDate") Long endDate);

    //获取客户每日第一次催单数量
    List<RPTCustomerReminderEntity> getReminderFirstQtyList(@Param("beginDate") Long beginDate,
                                                            @Param("endDate") Long endDate);

    //获取客户每日多次催单数量
    List<RPTCustomerReminderEntity> getReminderMultipleQtyList(@Param("beginDate") Long beginDate,
                                                               @Param("endDate") Long endDate);

    //获取客户催单的工单数量
    List<RPTCustomerReminderEntity> getReminderOrderQtyList(@Param("beginDate") Date beginDate,
                                                            @Param("endDate") Date endDate);

    //获取客户催单的工单数量
    List<RPTCustomerReminderEntity> getExceed48hourReminderQtyList(@Param("beginDate") Long beginDate,
                                                                   @Param("endDate") Long endDate);

    //获取客户24小时完成的催单数量
    List<RPTCustomerReminderEntity> getComplete24hourQtyList(@Param("beginDate") Long beginDate,
                                                             @Param("endDate") Long endDate);

    //获取客户超过48小时催单 24小时完成的催单数量
    List<RPTCustomerReminderEntity> getOver48ReminderCompletedQtyList(@Param("beginDate") Long beginDate,
                                                                      @Param("endDate") Long endDate);

    void insertCustomerReminder(RPTCustomerReminderEntity entity);

    void deleteCustomerReminders(@Param("systemId") Integer systemId,
                                 @Param("beginDate") Long beginDate,
                                 @Param("endDate") Long endDate);


    List<RPTCustomerReminderEntity> getCustomerReminderIds(@Param("systemId") Integer systemId,
                                                           @Param("beginDate") Long beginDate,
                                                           @Param("endDate") Long endDate,
                                                           @Param("quarter") String quarter);

    List<RPTCustomerReminderEntity> getCustomerReminderList(@Param("customerId") Long customerId,
                                                            @Param("systemId") Integer systemId,
                                                            @Param("productCategoryIds") List<Long> productCategoryIds,
                                                            @Param("beginDate") Long beginDate,
                                                            @Param("endDate") Long endDate,
                                                            @Param("quarter") String quarter);

    Integer hasReportData(@Param("customerId") Long customerId,
                          @Param("systemId") Integer systemId,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("beginDate") Long beginDate,
                          @Param("endDate") Long endDate,
                          @Param("quarter") String quarter);
}
