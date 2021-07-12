package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CompletedOrderRptMapper {

    //region 从Web数据库读数据

    List<RPTCompletedOrderEntity> getCompletedOrderListFromWebDB(@Param("beginDate") Date beginDate,
                                                                 @Param("endDate") Date endDate);

    List<RPTOrderDetail> getCompletedOrderDetailListFromWebDB(@Param("beginDate") Date beginDate,
                                                              @Param("endDate") Date endDate);

    List<RPTOrderItem> getCompletedOrderUnitBarcodeListFromWebDB(@Param("beginDate") Date beginDate,
                                                                 @Param("endDate") Date endDate);

    List<RPTOrderDetail> getCompletedOrderDetailFeeListWebDB(@Param("beginDate") Date beginDate,
                                                              @Param("endDate") Date endDate);
    //endregion 从Web数据库读数据


    //region 操作中间表

    /**
     * 插入完工单到中间表
     */
    void insert(RPTCompletedOrderEntity entity);

    Page<RPTCompletedOrderEntity> getCompletedOrderListByPaging(RPTCompletedOrderSearch search);

    List<LongTwoTuple> getCompletedOrderIds(@Param("quarter") String quarter,
                                            @Param("systemId") Integer systemId,
                                            @Param("beginChargeDate") Long beginChargeDate,
                                            @Param("endChargeDate") Long endChargeDate);

    void deleteCompletedOrders(@Param("quarter") String quarter,
                               @Param("systemId") Integer systemId,
                               @Param("beginChargeDate") Long beginChargeDate,
                               @Param("endChargeDate") Long endChargeDate);

    List<RPTCompletedOrderEntity> getCompletedOrderList(@Param("systemId") Integer systemId,
                                                        @Param("beginDate") Long beginDate,
                                                        @Param("endDate") Long endDate,
                                                        @Param("customerId") Long customerId,
                                                        @Param("quarter") String quarter);


    Integer hasReportData(RPTCompletedOrderSearch condition);
    //endregion 操作中间表
}
