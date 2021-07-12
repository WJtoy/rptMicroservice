package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;


@Mapper
public interface CustomerChargeSummaryRptMapper {

    //region 查询所有客户的工单数量

    List<RPTCustomerChargeSummaryMonthlyEntity> getNewOrderQtyList(@Param("quarter") String quarter,
                                                                   @Param("beginDate") Date beginDate,
                                                                   @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getCompletedOrderQtyList(@Param("quarter") String quarter,
                                                                         @Param("beginDate") Date beginDate,
                                                                         @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getReturnedOrderQtyList(@Param("beginDate") Date beginDate,
                                                                        @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getCancelledOrderQtyList(@Param("beginDate") Date beginDate,
                                                                         @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getUnCompletedOrderQtyList(@Param("endDate") Date endDate);

    //endregion 查询所有客户的工单数量

    //region 查询单个客户的工单数量

    Integer getNewOrderQty(@Param("quarter") String quarter,
                           @Param("customerId") Long customerId,
                           @Param("beginDate") Date beginDate,
                           @Param("endDate") Date endDate);

    Integer getCompletedOrderQty(@Param("quarter") String quarter,
                                 @Param("customerId") Long customerId,
                                 @Param("beginDate") Date beginDate,
                                 @Param("endDate") Date endDate);

    Integer getReturnedOrderQty(@Param("customerId") Long customerId,
                                @Param("beginDate") Date beginDate,
                                @Param("endDate") Date endDate);

    Integer getCancelledOrderQty(@Param("customerId") Long customerId,
                                 @Param("beginDate") Date beginDate,
                                 @Param("endDate") Date endDate);

    Integer getUnCompletedOrderQty(@Param("customerId") Long customerId,
                                   @Param("endDate") Date endDate);
    //endregion 查询单个客户的工单数量

    //region 查询所有客户的消费金额

    List<RPTCustomerChargeSummaryMonthlyEntity> getRechargeAmountList(@Param("quarter") String quarter,
                                                                      @Param("beginDate") Date beginDate,
                                                                      @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getCompletedOrderAndTimelinessAndUrgentChargeList(@Param("quarter") String quarter,
                                                                                                  @Param("beginDate") Date beginDate,
                                                                                                  @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getWriteOffChargeList(@Param("quarter") String quarter,
                                                                      @Param("beginDate") Date beginDate,
                                                                      @Param("endDate") Date endDate);

    List<RPTCustomerChargeSummaryMonthlyEntity> getBalanceList(@Param("endDate") Date endDate);

//    List<RPTCustomerChargeSummaryMonthlyEntity> getBlockAmountList(@Param("endDate") Date endDate);

    //endregion 查询所有客户的消费金额

    //region 查询单个客户的消费金额

    Double getRechargeAmount(@Param("quarter") String quarter,
                             @Param("customerId") Long customerId,
                             @Param("beginDate") Date beginDate,
                             @Param("endDate") Date endDate);

    RPTCustomerChargeSummaryMonthlyEntity getCompletedOrderAndTimelinessAndUrgentCharge(@Param("quarter") String quarter,
                                                                                        @Param("customerId") Long customerId,
                                                                                        @Param("beginDate") Date beginDate,
                                                                                        @Param("endDate") Date endDate);

    Double getWriteOffCharge(@Param("quarter") String quarter,
                             @Param("customerId") Long customerId,
                             @Param("beginDate") Date beginDate,
                             @Param("endDate") Date endDate);

    Double getBalance(@Param("customerId") Long customerId,
                      @Param("endDate") Date endDate);

//    Double getBlockAmount(@Param("customerId") Long customerId,
//                          @Param("endDate") Date endDate);

//    Double getCurrentMonthBlockAmount(@Param("quarter") String quarter,
//                                      @Param("customerId") Long customerId,
//                                      @Param("beginDate") Date beginDate,
//                                      @Param("endDate") Date endDate);


    //endregion 查询单个客户的消费金额

    RPTCustomerChargeSummaryMonthlyEntity getCustomerFinanceMonthly(@Param("systemId") Integer systemId,
                                                                    @Param("customerId") Long customerId,
                                                                    @Param("yearmonth") Integer yearmonth);

    List<LongTwoTuple> getCustomerFinanceMonthlyIds(@Param("systemId") Integer systemId,
                                                    @Param("yearmonth") Integer yearmonth);

    void deleteCustomerFinanceMonthly(@Param("systemId") Integer systemId,
                                      @Param("yearmonth") Integer yearmonth);

    void insertCustomerFinanceMonthly(RPTCustomerChargeSummaryMonthlyEntity entity);

    void updateCustomerFinanceMonthly(RPTCustomerChargeSummaryMonthlyEntity entity);

    RPTCustomerChargeSummaryMonthlyEntity getCustomerOrderQtyMonthly(@Param("systemId") Integer systemId,
                                                                     @Param("customerId") Long customerId,
                                                                     @Param("yearmonth") Integer yearmonth);

    List<LongTwoTuple> getCustomerOrderQtyMonthlyIds(@Param("systemId") Integer systemId,
                                                     @Param("yearmonth") Integer yearmonth);

    void deleteCustomerOrderQtyMonthly(@Param("systemId") Integer systemId,
                                       @Param("yearmonth") Integer yearmonth);

    void insertCustomerOrderQtyMonthly(RPTCustomerChargeSummaryMonthlyEntity entity);

    void updateCustomerOrderQtyMonthly(RPTCustomerChargeSummaryMonthlyEntity entity);

}
