package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

@Mapper
public interface CustomerReminderChartMapper {

    RPTCustomerReminderEntity getCustomerReminderList(@Param("systemId") Integer systemId,
                                                      @Param("beginDate") Long beginDate,
                                                      @Param("endDate") Long endDate,
                                                      @Param("quarter") String quarter);
}
