package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.provider.rpt.entity.ComplainDailyEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ComplainRatioDailyRptMapper {

    List<ComplainDailyEntity> getComplainDailyList(@Param("startDate") Date startDate,
                                                   @Param("endDate") Date endDate);


    List<LongTwoTuple> getComplainOrderIds(@Param("systemId") Integer systemId,
                                           @Param("startDate") Long startDate,
                                           @Param("endDate") Long endDate,
                                           @Param("quarter") String quarter);

    void insertComplainDaily(ComplainDailyEntity entity);

    void updateComplainDaily(ComplainDailyEntity entity);


    void deleteComplainData(@Param("startDate") Long startDate,
                            @Param("endDate") Long endDate,
                            @Param("systemId") Integer systemId,
                            @Param("quarter") String quarter);

    List<RPTAreaCompletedDailyEntity> getProvinceComplainOrderData(RPTGradedOrderSearch search);

    List<RPTAreaCompletedDailyEntity> getProvinceServicePointBadOrderData(RPTGradedOrderSearch search);

    List<RPTAreaCompletedDailyEntity> getCityComplainOrderData(RPTGradedOrderSearch search);

    List<RPTAreaCompletedDailyEntity> getCityServicePointBadOrderData(RPTGradedOrderSearch search);

    List<RPTAreaCompletedDailyEntity> getAreaComplainOrderData(RPTGradedOrderSearch search);

    List<RPTAreaCompletedDailyEntity> getAreaServicePointBadOrderData(RPTGradedOrderSearch search);

    Integer hasAreaComplainCompletedReportData(RPTGradedOrderSearch search);

}
