package com.kkl.kklplus.provider.rpt.mapper;


import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteEntity;
import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteSearch;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

/**
 * @Auther wj
 * @Date 2021/5/24 10:45
 */
@Mapper
public interface CustomerEveryDayCompleteMapper {


    /**
     * 省每日下单数量
     * @param search
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasProvinceCompleteOrderData(RPTEveryDayCompleteSearch search);

    /**
     * 市每日下单数量
     * @param search
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasCompleteOrderData(RPTEveryDayCompleteSearch search);

    /**
     * 完工数量
     * @param search
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasCompleteRateData(RPTEveryDayCompleteSearch search);

    /**
     * 72时效完工数
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasCompleteRate72Data(@Param("beforeEndDate") Date beforeEndDate,
                                                            @Param("startDate") Date startDate,
                                                          @Param("endDate") Date endDate,
                                                          @Param("areaType") Integer areaType,
                                                          @Param("areaId")Long areaId,
                                                          @Param("customerId") Long customerId,
                                                          @Param("quarters") List<String> quarters);

    /**
     * 到货48小时完工数
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasArrivalCompleteRateData(@Param("beforeEndDate") Date beforeEndDate,
                                                                @Param("startDate") Date startDate,
                                                               @Param("endDate") Date endDate,
                                                               @Param("areaType") Integer areaType,
                                                               @Param("areaId")Long areaId,
                                                               @Param("customerId") Long customerId,
                                                               @Param("quarters") List<String> quarters);

    /**
     * 未完工单
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> unCompleteOrderData(
                                                        @Param("endDate") Date endDate,
                                                        @Param("areaType") Integer areaType,
                                                        @Param("areaId")Long areaId,
                                                        @Param("customerId") Long customerId,
                                                        @Param("quarters") List<String> quarters);


    /**
     * 到货72小时下单
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarter
     * @return
     */
    List<RPTEveryDayCompleteEntity> arrivalOrderData72(
                                                          @Param("startDate") Date startDate,
                                                          @Param("endDate") Date endDate,
                                                          @Param("areaType") Integer areaType,
                                                          @Param("areaId")Long areaId,
                                                          @Param("customerId") Long customerId,
                                                          @Param("quarter") String quarter);

    /**
     * 无到货72小时下单
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarter
     * @return
     */
    List<RPTEveryDayCompleteEntity> unArrivalOrderData72(
                                                       @Param("startDate") Date startDate,
                                                       @Param("endDate") Date endDate,
                                                       @Param("areaType") Integer areaType,
                                                       @Param("areaId")Long areaId,
                                                       @Param("customerId") Long customerId,
                                                       @Param("quarter") String quarter);


    /**
     * 周  到货下单
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarter
     * @return
     */
    List<RPTEveryDayCompleteEntity> arrivalOrderDataWeek(
                                                        @Param("startDate") Date startDate,
                                                        @Param("endDate") Date endDate,
                                                        @Param("areaType") Integer areaType,
                                                        @Param("areaId")Long areaId,
                                                        @Param("customerId") Long customerId,
                                                        @Param("quarter") String quarter);



    /**
     * 周  无到货下单
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarter
     * @return
     */
    List<RPTEveryDayCompleteEntity> unArrivalOrderDataWeek(
                                                         @Param("startDate") Date startDate,
                                                         @Param("endDate") Date endDate,
                                                         @Param("areaType") Integer areaType,
                                                         @Param("areaId")Long areaId,
                                                         @Param("customerId") Long customerId,
                                                         @Param("quarter") String quarter);


    /**
     * 有到货72小时完工单
     * @param beforeEndDate
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasArrivalCompleteRateData72(@Param("beforeEndDate") Date beforeEndDate,
                                                         @Param("startDate") Date startDate,
                                                         @Param("endDate") Date endDate,
                                                         @Param("areaType") Integer areaType,
                                                         @Param("areaId")Long areaId,
                                                         @Param("customerId") Long customerId,
                                                         @Param("quarters") List<String> quarters);


    /**
     * 无到货72小时完工单
     * @param beforeEndDate
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> unArrivalCompleteRateData72(@Param("beforeEndDate") Date beforeEndDate,
                                                                 @Param("startDate") Date startDate,
                                                                 @Param("endDate") Date endDate,
                                                                 @Param("areaType") Integer areaType,
                                                                 @Param("areaId")Long areaId,
                                                                 @Param("customerId") Long customerId,
                                                                 @Param("quarters") List<String> quarters);

    /**
     * 无到货 周  完工单
     * @param beforeEndDate
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> unArrivalCompleteRateDataWeek(@Param("beforeEndDate") Date beforeEndDate,
                                                                @Param("startDate") Date startDate,
                                                                @Param("endDate") Date endDate,
                                                                @Param("areaType") Integer areaType,
                                                                @Param("areaId")Long areaId,
                                                                @Param("customerId") Long customerId,
                                                                @Param("quarters") List<String> quarters);


    /**
     * 有到货 周  完工单
     * @param beforeEndDate
     * @param startDate
     * @param endDate
     * @param areaType
     * @param areaId
     * @param customerId
     * @param quarters
     * @return
     */
    List<RPTEveryDayCompleteEntity> hasArrivalCompleteRateDataWeek(@Param("beforeEndDate") Date beforeEndDate,
                                                                  @Param("startDate") Date startDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("areaType") Integer areaType,
                                                                  @Param("areaId")Long areaId,
                                                                  @Param("customerId") Long customerId,
                                                                  @Param("quarters") List<String> quarters);

    /**
     * 检查是否有数据可以导出
     * @param search
     * @return
     */
    Integer hasReportData(RPTEveryDayCompleteSearch search);
}
