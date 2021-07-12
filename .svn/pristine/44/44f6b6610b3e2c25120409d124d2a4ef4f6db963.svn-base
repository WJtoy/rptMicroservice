package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTUncompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTUncompletedOrderSearch;
import org.apache.ibatis.annotations.Mapper;


@Mapper
public interface UncompletedOrderRptMapper {

    /**
     * 查询客户的未完工的工单信息
     * @return
     */
    Page<RPTUncompletedOrderEntity> getUncompletedOrder(RPTUncompletedOrderSearch search);


    Integer hasReportData(RPTUncompletedOrderSearch condition);
}
