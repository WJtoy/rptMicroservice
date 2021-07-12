package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.entity.common.NameValuePair;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.math.BigDecimal;

@Mapper
public interface IncurExpenseChartMapper {

     NameValuePair<BigDecimal,BigDecimal> getSpecialExpenses(@Param("systemId") Integer systemId,
                                                             @Param("yearmonth") String yearmonth,
                                                             @Param("day") Integer day);
}
