package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface CompletedOrderDetailsRptMapper {


    Integer getCompletedOrderSum(RPTCompletedOrderDetailsSearch search);


    Integer getCustomerWriteOffSum(RPTCompletedOrderDetailsSearch search);


    Integer getServicePointWriteOffSum(RPTCompletedOrderDetailsSearch search);

    Page<RPTCompletedOrderDetailsEntity> getCompletedOrderList(RPTCompletedOrderDetailsSearch search);

    Page<RPTCompletedOrderDetailsEntity> getCustomerWriteOffList(RPTCompletedOrderDetailsSearch search);

    Page<RPTCompletedOrderDetailsEntity> getServicePointWriteOffList(RPTCompletedOrderDetailsSearch search);


}
