package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTOrderDailyWorkEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface RptCreatedOrderMapper {

    List<RPTOrderDailyWorkEntity>  getOrderDailyWorkList(RPTCompletedOrderDetailsSearch search);

    Integer getOrderDailyWorkSum(RPTCompletedOrderDetailsSearch search);

}
