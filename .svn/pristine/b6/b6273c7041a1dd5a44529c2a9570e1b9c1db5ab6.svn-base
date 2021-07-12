package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTDispatchOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface DispatchListRptMapper {


    List<RPTDispatchOrderEntity> getPlanInformationSum(RPTCompletedOrderDetailsSearch search);


    List<RPTDispatchOrderEntity> getCancelOrderSum(RPTCompletedOrderDetailsSearch search);


    List<RPTDispatchOrderEntity> getTheTotalOrder(RPTCompletedOrderDetailsSearch search);

    Integer hasReportData(RPTCompletedOrderDetailsSearch search);



}
