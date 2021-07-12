package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointPaySummaryEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointPaySummarySearch;
import com.kkl.kklplus.provider.rpt.entity.RPTServicePointChargeEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ServicePointChargeRptMapper {
    /**
     * 获取网点应付汇总金额
     */
    Page<RPTServicePointPaySummaryEntity> getServicePointPaySummaryRptData(RPTServicePointPaySummarySearch search);

    /**
     * 获取网点成本
     */
    Page<RPTServicePointPaySummaryEntity> getServicePointCostPerRptData(RPTServicePointPaySummarySearch search);

    List<RPTServicePointPaySummaryEntity> getServicePointCostPerRptList(RPTServicePointPaySummarySearch search);

    Integer hasServicePointPaySummaryReportData(RPTServicePointPaySummarySearch search);

    Integer hasServicePointCostPerReportData(RPTServicePointPaySummarySearch search);

//    List<RPTServicePointPaySummaryEntity> getServicePointServiceAreas(@Param("servicePointIds") List<Long> servicePointIds);
//
//    List<RPTServicePointPaySummaryEntity> getServicePointServiceAreasNew(RPTServicePointPaySummarySearch search);

    /**
     * 网点的当月应付合计
     */
    List<RPTServicePointChargeEntity> getPayableChargeA(@Param("startDate") Date startDate,
                                                        @Param("endDate") Date endDate,
                                                        @Param("quarter") String quarter,
                                                        @Param("beginLimit") Integer beginLimit,
                                                        @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getPayableChargeB(@Param("startDate") Date startDate,
                                                        @Param("endDate") Date endDate,
                                                        @Param("quarter") String quarter,
                                                        @Param("beginLimit") Integer beginLimit,
                                                        @Param("endLimit") Integer endLimit);

    /**
     * 网点的当月已付合计
     */
    List<RPTServicePointChargeEntity> getServicePointTotalPaid(@Param("selectedYear") Integer selectedYear,
                                                               @Param("selectedMonth") Integer selectedMonth,
                                                               @Param("beginLimit") Integer beginLimit,
                                                               @Param("endLimit") Integer endLimit);

    /**
     * 网点的当月余额合计
     */
    List<RPTServicePointChargeEntity> getServicePointTotalCurrBalance(@Param("selectedYear") Integer selectedYear,
                                                                      @Param("selectedMonth") Integer selectedMonth,
                                                                      @Param("beginLimit") Integer beginLimit,
                                                                      @Param("endLimit") Integer endLimit);

    /**
     * 网点完工单总金额
     */
    List<RPTServicePointChargeEntity> getServicePointCompletedOrderCharge(@Param("beginDate") Date beginDate,
                                                                          @Param("endDate") Date endDate,
                                                                          @Param("quarter") String quarter,
                                                                          @Param("beginLimit") Integer beginLimit,
                                                                          @Param("endLimit") Integer endLimit);

    /**
     * 网点的时效费用、保险费用
     */
    List<RPTServicePointChargeEntity> getServicePointTimelinessAndInsurance(@Param("beginDate") Date beginDate,
                                                                            @Param("endDate") Date endDate,
                                                                            @Param("quarter") String quarter,
                                                                            @Param("beginLimit") Integer beginLimit,
                                                                            @Param("endLimit") Integer endLimit);

    /**
     * 网点退补总金额
     */
    List<RPTServicePointChargeEntity> getServicePointDiffCharge(@Param("beginDate") Date beginDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("quarter") String quarter,
                                                                @Param("beginLimit") Integer beginLimit,
                                                                @Param("endLimit") Integer endLimit);

    /**
     * 网点的指定月余额合计
     */
    List<RPTServicePointChargeEntity> getServicePointTotalBalance(@Param("selectedYear") Integer selectedYear,
                                                                  @Param("selectedMonth") Integer selectedMonth,
                                                                  @Param("beginLimit") Integer beginLimit,
                                                                  @Param("endLimit") Integer endLimit);

    /**
     * 获取网点的其他费用和远程费用
     */
    List<RPTServicePointChargeEntity> getServicePointOtherChargeAndTravelCharge(@Param("beginDate") Date beginDate,
                                                                                @Param("endDate") Date endDate,
                                                                                @Param("quarter") String quarter,
                                                                                @Param("beginLimit") Integer beginLimit,
                                                                                @Param("endLimit") Integer endLimit);

    /**
     * 查询网点的完工单数量
     */
    List<RPTServicePointChargeEntity> getAllServicePointFinishQty(@Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("beginLimit") Integer beginLimit,
                                                                  @Param("endLimit") Integer endLimit);

    /**
     * 查询网点的平台服务费
     */
    List<RPTServicePointChargeEntity> getAllServicePointPlatFees(@Param("quarter") String quarter,
                                                                 @Param("beginDate") Date beginDate,
                                                                 @Param("endDate") Date endDate,
                                                                 @Param("beginLimit") Integer beginLimit,
                                                                 @Param("endLimit") Integer endLimit);

    List<RPTServicePointChargeEntity> getUpServicePointChargeData(@Param("systemId") Integer systemId,
                                                                  @Param("yearMonth") Integer yearMonth,
                                                                  @Param("quarter") String quarter);

    List<RPTServicePointChargeEntity> getUpdatePlatformFeeData(@Param("systemId") Integer systemId,
                                                               @Param("yearMonth") Integer yearMonth,
                                                               @Param("quarter") String quarter);

    int updateServicePointCharge(RPTServicePointChargeEntity entity);

    int updateServicePointChargeInfoFee(RPTServicePointChargeEntity entity);

    void updatePlatformFee(RPTServicePointChargeEntity entity);

    void insertServicePointCharge(RPTServicePointChargeEntity entity);

    void deleteServicePointCharge(@Param("systemId") Integer systemId,
                                  @Param("yearMonth") Integer yearMonth,
                                  @Param("quarter") String quarter);


    List<RPTServicePointChargeEntity> getUpdateServicePointQuarter(@Param("systemId") Integer systemId,
                                                                   @Param("yearMonth") Integer yearMonth,
                                                                   @Param("quarter") String quarter);

    int insertServicePointChargeNew(RPTServicePointChargeEntity entity);

    List<RPTServicePointChargeEntity> getUpdatePlatformFeeList(@Param("systemId") Integer systemId,
                                                               @Param("yearMonth") Integer yearMonth,
                                                               @Param("quarter") String quarter);
}
