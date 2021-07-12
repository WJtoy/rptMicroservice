package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.provider.rpt.entity.RPTServicePointChargeEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ServicePointChargeNewRptMapper {

    //region 读取Web表数据

    List<RPTServicePointChargeEntity> getPayableAList(@Param("startDate") Date startDate,
                                                      @Param("endDate") Date endDate,
                                                      @Param("quarter") String quarter,
                                                      @Param("beginLimit") Integer beginLimit,
                                                      @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getPayableBList(@Param("startDate") Date startDate,
                                                      @Param("endDate") Date endDate,
                                                      @Param("quarter") String quarter,
                                                      @Param("beginLimit") Integer beginLimit,
                                                      @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getPaidAmountList(@Param("selectedYear") Integer selectedYear,
                                                        @Param("selectedMonth") Integer selectedMonth,
                                                        @Param("beginLimit") Integer beginLimit,
                                                        @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getLastMonthBalanceList(@Param("selectedYear") Integer selectedYear,
                                                              @Param("selectedMonth") Integer selectedMonth,
                                                              @Param("beginLimit") Integer beginLimit,
                                                              @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getTheBalanceList(@Param("selectedYear") Integer selectedYear,
                                                        @Param("selectedMonth") Integer selectedMonth,
                                                        @Param("beginLimit") Integer beginLimit,
                                                        @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getCompletedChargeList(@Param("beginDate") Date beginDate,
                                                             @Param("endDate") Date endDate,
                                                             @Param("quarter") String quarter,
                                                             @Param("beginLimit") Integer beginLimit,
                                                             @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getReturnChargeList(@Param("beginDate") Date beginDate,
                                                          @Param("endDate") Date endDate,
                                                          @Param("quarter") String quarter,
                                                          @Param("beginLimit") Integer beginLimit,
                                                          @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getTimelineUrgentInsurancePraiseTaxInfoList(@Param("beginDate") Date beginDate,
                                                                                  @Param("endDate") Date endDate,
                                                                                  @Param("quarter") String quarter,
                                                                                  @Param("beginLimit") Integer beginLimit,
                                                                                  @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getPlatformFeeList(@Param("quarter") String quarter,
                                                         @Param("beginDate") Date beginDate,
                                                         @Param("endDate") Date endDate,
                                                         @Param("beginLimit") Integer beginLimit,
                                                         @Param("endLimit") Integer endLimit);
    /**
     * 查询网点质保金(财务充值)
     */
    List<RPTServicePointChargeEntity> getAllRechargeDepositFees(@Param("quarter") String quarter,
                                                                @Param("beginDate") Date beginDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("beginLimit") Integer beginLimit,
                                                                @Param("endLimit") Integer endLimit);

    /**
     * 查询网点质保金(订单扣款)
     */
    List<RPTServicePointChargeEntity> getAllEngineerDepositFees(@Param("quarter") String quarter,
                                                                @Param("beginDate") Date beginDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("beginLimit") Integer beginLimit,
                                                                @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getOtherTravelChargeList(@Param("beginDate") Date beginDate,
                                                               @Param("endDate") Date endDate,
                                                               @Param("quarter") String quarter,
                                                               @Param("beginLimit") Integer beginLimit,
                                                               @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getCompleteQtyList(@Param("beginDate") Date beginDate,
                                                         @Param("endDate") Date endDate,
                                                         @Param("beginLimit") Integer beginLimit,
                                                         @Param("endLimit") Integer endLimit);

    //endregion 读取Web表数据

    void insert(RPTServicePointChargeEntity entity);

    void delete(@Param("quarter") String quarter,
                @Param("systemId") Integer systemId,
                @Param("yearMonth") Integer yearMonth);


    List<RPTServicePointChargeEntity> getServicePointChargeIds(@Param("quarter") String quarter,
                                                               @Param("systemId") Integer systemId,
                                                               @Param("yearMonth") Integer yearMonth,
                                                               @Param("beginLimit") Integer beginLimit,
                                                               @Param("endLimit") Integer endLimit);


    int update(RPTServicePointChargeEntity entity);

}
