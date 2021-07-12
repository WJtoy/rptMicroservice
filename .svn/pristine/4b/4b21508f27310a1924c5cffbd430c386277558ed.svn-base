package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointOderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointCompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.entity.rpt.web.RPTOrderServicePointFee;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ServicePointCompletedOrderRptMapper {

    //region 从Web数据库读数据

    List<RPTServicePointCompletedOrderEntity> getServicePointCompletedOrderListFromWebDB(@Param("beginDate") Date beginDate,
                                                                                         @Param("endDate") Date endDate);

    List<RPTOrderServicePointFee> getOrderServicePointFeeListFromWebDB(@Param("beginDate") Date beginDate,
                                                                       @Param("endDate") Date endDate);

    List<RPTOrderDetail> getServicePointCompletedOrderDetailListFromWebDB(@Param("beginDate") Date beginDate,
                                                                          @Param("endDate") Date endDate);

    //endregion 从Web数据库读数据


    //region 操作中间表

    /**
     * 插入完工单到中间表
     */
    void insert(RPTServicePointCompletedOrderEntity entity);

    List<LongTwoTuple> getServicePointCompletedOrderIds(@Param("quarter") String quarter,
                                                        @Param("systemId") Integer systemId,
                                                        @Param("beginChargeDate") Long beginChargeDate,
                                                        @Param("endChargeDate") Long endChargeDate);

    void deleteServicePointCompletedOrders(@Param("quarter") String quarter,
                                           @Param("systemId") Integer systemId,
                                           @Param("beginChargeDate") Long beginChargeDate,
                                           @Param("endChargeDate") Long endChargeDate);

    Page<RPTServicePointCompletedOrderEntity> getServicePointCompletedOrderListByPaging(RPTServicePointCompletedOrderSearch search);

    List<RPTServicePointOderEntity> getCompletedOrderByPaging(RPTServicePointWriteOffSearch search);

    Integer getPointCompletedOrderSum(RPTServicePointWriteOffSearch search);



    //endregion 操作中间表
}
