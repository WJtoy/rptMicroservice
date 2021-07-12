package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTUncompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTUncompletedQtyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTUncompletedOrderSearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface UncompletedQtyRptMapper {


    List<RPTUncompletedQtyEntity> getUncompletedQtyList(RPTUncompletedOrderSearch search);

    List<RPTUncompletedOrderEntity> getUncompletedOrderList(RPTUncompletedOrderSearch search);


    Integer hasReportData(RPTUncompletedOrderSearch condition);
}
