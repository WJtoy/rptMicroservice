package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTDevelopAverageOrderFeeEntity;
import com.kkl.kklplus.entity.rpt.RPTGradedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTKefuCompletedDailyEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface GradedOrderRptMapper {

    /**
     * 查询原表中的客评工单
     * @param beginDate
     * @param endDate
     * @return
     */
     List<RPTGradedOrderEntity> getGradedOrderData(@Param("beginDate")Date beginDate,
                                                         @Param("endDate")Date endDate
                                                         );

    /**
     *获取客评人
     */
    List<TwoTuple<Long,Long>> getGradeCreateByList(@Param("beginDate")Date beginDate,
                                        @Param("endDate")Date endDate
    );


    /**
     * 根据队列查询原表中的客评工单
     * @param quarter
     * @param orderId
     * @return
     */
    RPTGradedOrderEntity getGradedOrderDataOfMQ(@Param("quarter")String quarter,
                                                  @Param("orderId")Long orderId
    );

    /**
     * 查询原表中突击单数据
     * @param beginDate
     * @param endDate
     * @return
     */
     List<RPTGradedOrderEntity> getCrushOrderData(@Param("beginDate")Date beginDate,
                                                  @Param("endDate")Date endDate);

    /**
     * 根据消息队列查询原表中突击单数据
     * @param quarter
     * @param orderId
     * @return
     */
    RPTGradedOrderEntity getCrushOrderDataOfMQ(@Param("quarter")String quarter,
                                                 @Param("orderId")Long orderId);

    /**
     *获取客评人
     */
    List<TwoTuple<Long,Long>> getGradeCreateByOfMQ(@Param("quarter")String quarter,
                                                   @Param("orderId")Long orderId
    );

    /**
     * 写入中间表
     * @param entity
     */
     void insertGradedOrder(RPTGradedOrderEntity entity);

    /**
     *查询中间表已有的客评工单id
     * @return
     */
    List<LongTwoTuple> getGradedOrderIds(@Param("quarter") String quarter,
                                            @Param("systemId") Integer systemId,
                                            @Param("beginDate") Long beginDate,
                                            @Param("endDate") Long endDate);
    /**
     *查询中间表中重复的订单id
     * @return
     */
    List<LongTwoTuple> getHavingGradedOrder(@Param("quarter") String quarter,
                                         @Param("systemId") Integer systemId,
                                         @Param("beginDate") Long beginDate,
                                         @Param("endDate") Long endDate);

    void deleteHavingGradedOrder(@Param("orderId")Long orderId,
                                 @Param("count")Integer count
                                 );



    void deleteGradedOrders(@Param("quarter") String quarter,
                               @Param("systemId") Integer systemId,
                               @Param("beginDate") Long beginDate,
                               @Param("endDate") Long endDate);


    /**
     * 获取省每日完工报表数据
     * @param beginDate
     * @param endDate
     * @param systemId
     * @param areaId
     * @param areaType
     * @param customerId
     * @param quarter
     * @param productCategoryIds
     * @return
     */
     List<RPTAreaCompletedDailyEntity> getProvinceGradedOrderData(@Param("beginDate")Long beginDate,
                                                                  @Param("endDate")Long endDate,
                                                                  @Param("systemId")Integer systemId,
                                                                  @Param("areaId")Long areaId,
                                                                  @Param("areaType")Integer areaType,
                                                                  @Param("customerId")Long customerId,
                                                                  @Param("quarter")String quarter,
                                                                  @Param("productCategoryIds")List<Long> productCategoryIds
                                                       );

    /**
     * 获取市每日完工报表数据
     * @param beginDate
     * @param endDate
     * @param systemId
     * @param areaId
     * @param areaType
     * @param customerId
     * @param quarter
     * @param productCategoryIds
     * @return
     */
    List<RPTAreaCompletedDailyEntity> getCityGradedOrderData(@Param("beginDate")Long beginDate,
                                                          @Param("endDate")Long endDate,
                                                          @Param("systemId")Integer systemId,
                                                          @Param("areaId")Long areaId,
                                                          @Param("areaType")Integer areaType,
                                                          @Param("customerId")Long customerId,
                                                          @Param("quarter")String quarter,
                                                          @Param("productCategoryIds")List<Long> productCategoryIds
    );

    /**
     * 获取区每日完工报表数据
     * @param beginDate
     * @param endDate
     * @param systemId
     * @param areaId
     * @param areaType
     * @param customerId
     * @param quarter
     * @param productCategoryIds
     * @return
     */
    List<RPTAreaCompletedDailyEntity> getAreaGradedOrderData(@Param("beginDate")Long beginDate,
                                                             @Param("endDate")Long endDate,
                                                             @Param("systemId")Integer systemId,
                                                             @Param("areaId")Long areaId,
                                                             @Param("areaType")Integer areaType,
                                                             @Param("customerId")Long customerId,
                                                             @Param("quarter")String quarter,
                                                             @Param("productCategoryIds")List<Long> productCategoryIds
    );

    /**
     * 获取客服每日完工报表数据
     * @param beginDate
     * @param endDate
     * @param kefuId
     * @param quarter
     * @return
     */
    List<RPTKefuCompletedDailyEntity> getKefuGradedOrderData(@Param("beginDate")Long beginDate,
                                                             @Param("endDate")Long endDate,
                                                             @Param("systemId")Integer systemId,
                                                             @Param("kefuId")Long kefuId,
                                                             @Param("productCategoryIds") List<Long> productCategoryIds,
                                                             @Param("quarter")String quarter
    );

    /**
     * 获取rpt_graded_order工单费用报表数据
     * @param beginDate
     * @param endDate
     * @param quarters
     * @param productCategoryIds
     * @return
     */
    Page<RPTGradedOrderEntity> getOrderServicePointFeeGradedOrderData(@Param("beginDate")Long beginDate,
                                                                      @Param("endDate")Long endDate,
                                                                      @Param("systemId")Integer systemId,
                                                                      @Param("servicePointId")Long servicePointId,
                                                                      @Param("quarters")List<String> quarters,
                                                                      @Param("productCategoryIds")List<Long> productCategoryIds);

    List<RPTDevelopAverageOrderFeeEntity> getDevelopAverageOrderFee(@Param("beginDate")Long beginDate,
                                                                    @Param("endDate")Long endDate,
                                                                    @Param("systemId")Integer systemId,
                                                                    @Param("quarter")String quarter,
                                                                    @Param("productCategoryIds")List<Long> productCategoryIds
                              );

    /**
     * 查询上个月的完工单总数
     */
    Long getCompletedOrderSum(@Param("beginDate")Long beginDate,
                              @Param("endDate")Long endDate,
                              @Param("quarter")String quarter,
                              @Param("systemId")Integer systemId,
                              @Param("productCategoryIds") List<Long> productCategoryIds,
                              @Param("keFuIds") List<Long> keFuIds
                              );


    Integer hasAreaCompletedOrderReportData(@Param("beginDate") Long beginDate,
                                            @Param("endDate") Long endDate,
                                            @Param("areaType") Integer areaType,
                                            @Param("areaId") Long areaId,
                                            @Param("systemId") Integer systemId,
                                            @Param("customerId") Long customerId,
                                            @Param("quarter") String quarter,
                                            @Param("productCategoryIds") List<Long> productCategoryIds,
                                            @Param("quarters") List<String> quarters
                                            );

    Integer hasKefuCompletedOrderReportData(@Param("beginDate") Long beginDate,
                                            @Param("endDate") Long endDate,
                                            @Param("kefuIds") List<Long> kefuIds,
                                            @Param("systemId") Integer systemId,
                                            @Param("productCategoryIds") List<Long> productCategoryIds,
                                            @Param("quarter")String quarter,
                                            @Param("quarters") List<String> quarters);


    Integer hasOrderServicePointFeeReportData(@Param("beginDate") Long beginDate,
                                              @Param("endDate") Long endDate,
                                              @Param("systemId") Integer systemId,
                                              @Param("servicePointId") Long servicePointId,
                                              @Param("quarter") String quarter,
                                              @Param("productCategoryIds") List<Long> productCategoryIds,
                                              @Param("quarters") List<String> quarters
    );

    Integer hasDevelopAverageFeeReportData(@Param("beginDate") Long beginDate,
                                           @Param("endDate") Long endDate,
                                           @Param("systemId") Integer systemId,
                                           @Param("quarter") String quarter,
                                           @Param("productCategoryIds") List<Long> productCategoryIds,
                                           @Param("quarters") List<String> quarters
    );
}
