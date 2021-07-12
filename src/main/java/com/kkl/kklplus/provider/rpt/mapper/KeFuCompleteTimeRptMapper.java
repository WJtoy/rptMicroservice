package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.provider.rpt.entity.KeFuCompleteTimeRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface KeFuCompleteTimeRptMapper {

    List<RPTKeFuCompleteTimeEntity> getKeFuCompleteTimeData(RPTKeFuCompleteTimeSearch search);

    List<RPTKeFuCompleteTimeEntity> getKeFuCompleteTimeDataNew(RPTKeFuCompleteTimeSearch search);

    Integer hasReportData(RPTKeFuCompleteTimeSearch search);

    Integer hasReportNewData(RPTKeFuCompleteTimeSearch search);

    List<KeFuCompleteTimeRptEntity> getOrderCreatedData(@Param("systemId") Integer systemId,
                                                        @Param("beginDate") Long beginDate,
                                                        @Param("endDate") Long endDate,
                                                        @Param("quarter") String quarter);

    List<KeFuCompleteTimeRptEntity> getOrderCloseData(@Param("systemId") Integer systemId,
                                                      @Param("beginDate") Long beginDate,
                                                      @Param("endDate") Long endDate,
                                                      @Param("quarter") String quarter);

    List<KeFuCompleteTimeRptEntity> getOrderCancelledData(@Param("systemId") Integer systemId,
                                                          @Param("beginDate") Long beginDate,
                                                          @Param("endDate") Long endDate,
                                                          @Param("quarter") String quarter);

    List<KeFuCompleteTimeRptEntity> getPlanTypeData(@Param("beginDate") Date beginDate,
                                                    @Param("endDate") Date endDate,
                                                    @Param("beginLimit") Integer beginLimit,
                                                    @Param("endLimit") Integer endLimit);

    List<KeFuCompleteTimeRptEntity> getComplainInformation(@Param("beginDate") Date beginDate,
                                                           @Param("endDate") Date endDate);

    List<LongTwoTuple> getNeedUpdateRptRecords(@Param("systemId") Integer systemId,
                                               @Param("beginDate") Long beginDate,
                                               @Param("endDate") Long endDate,
                                               @Param("quarter") String quarter);

    List<KeFuCompleteTimeRptEntity> getUpdateForRptServicePointId(@Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("quarter") String quarter);

    void insertOrderCreatedData(@Param("entityList") List<KeFuCompleteTimeRptEntity> entity);

    void updateOrderCloseData(KeFuCompleteTimeRptEntity entity);

    void updateOrderCancelledData(KeFuCompleteTimeRptEntity entity);


    void updateRPtPlanType(KeFuCompleteTimeRptEntity entity);

    void updateComplainInformation(KeFuCompleteTimeRptEntity entity);

    void updateOrderInformation(KeFuCompleteTimeRptEntity entity);

    void deleteOrderCreatedData(@Param("systemId") Integer systemId,
                                @Param("beginDate") Long beginDate,
                                @Param("endDate") Long endDate,
                                @Param("quarter") String quarter);

    void deleteComplainData(@Param("systemId") Integer systemId,
                            @Param("beginDate") Long beginDate,
                            @Param("endDate") Long endDate,
                            @Param("quarter") String quarter);

}
