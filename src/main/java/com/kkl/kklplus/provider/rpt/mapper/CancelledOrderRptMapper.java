package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCancelledOrderSearch;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;


@Mapper
public interface CancelledOrderRptMapper {

    List<RPTCancelledOrderEntity> getCancelledOrderListFromWebDB(@Param("beginDate") Date beginDate,
                                                                 @Param("endDate") Date endDate);

    RPTCancelledOrderEntity getCancelledOrderListFromWebMQ(@Param("orderId") Long orderId);

    void insert(RPTCancelledOrderEntity entity);

    Page<RPTCancelledOrderEntity> getCancelledOrderListByPaging(RPTCancelledOrderSearch search);

    Page<RPTCancelledOrderEntity> getCancelledOrderListNewByPaging(RPTCancelledOrderSearch search);


    List<LongTwoTuple> getCancelledOrderIds(@Param("quarter") String quarter,
                                            @Param("systemId") Integer systemId,
                                            @Param("beginCloseDate") Long beginCloseDate,
                                            @Param("endCloseDate") Long endCloseDate);

    void deleteCancelledOrders(@Param("quarter") String quarter,
                               @Param("systemId") Integer systemId,
                               @Param("beginCloseDate") Long beginCloseDate,
                               @Param("endCloseDate") Long endCloseDate);


    List<RPTCancelledOrderEntity> getCancelledOrderList(@Param("systemId") Integer systemId,
                                                        @Param("beginDate") Long beginDate,
                                                        @Param("endDate") Long endDate,
                                                        @Param("customerId") Long customerId,
                                                        @Param("quarter") String quarter);

    Integer hasReportData(RPTCancelledOrderSearch condition);

    Integer hasReportNewData(RPTCancelledOrderSearch condition);

}
