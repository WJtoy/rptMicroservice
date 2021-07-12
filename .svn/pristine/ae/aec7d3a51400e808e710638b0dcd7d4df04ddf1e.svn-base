package com.kkl.kklplus.provider.rpt.mapper;




import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTTravelChargeRankEntity;
import com.kkl.kklplus.entity.rpt.search.RPTTravelChargeRankSearchCondition;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.ServicePointServiceAreaEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;


@Mapper
public interface TravelChargeRankRptMapper {

     List<RPTTravelChargeRankEntity> getCompletedOrderCharge(@Param("beginDate") Date beginDate,
                                                             @Param("endDate") Date endDate,
                                                             @Param("quarter") String quarter
                                                             );


    List<RPTTravelChargeRankEntity> getWriteOffCharge(@Param("beginDate") Date beginDate,
                                                      @Param("endDate") Date endDate,
                                                      @Param("quarter") String quarter);


    List<RPTTravelChargeRankEntity> getTravelAndOtherCharge(@Param("beginDate") Date beginDate,
                                                    @Param("endDate") Date endDate,
                                                    @Param("quarter") String quarter);


    List<RPTTravelChargeRankEntity> getCompleteQty(@Param("beginDate") Date beginDate,
                                                    @Param("endDate") Date endDate,
                                                    @Param("quarter") String quarter);


//    List<ServicePointServiceAreaEntity> getServicePointServiceAreas(@Param("servicePointIds") List<Long> servicePointIds);

//    List<LongTwoTuple> getServicePointServiceAreaIds(@Param("servicePointIds") List<Long> servicePointIds);

    void insertTravelChargeRank(RPTTravelChargeRankEntity entity);

    void updateTravelChargeRank(RPTTravelChargeRankEntity entity);

    void deleteTravelChargeRankRpt(@Param("systemId") Integer systemId,
                                @Param("yearmonth") Integer yearmonth);

    List<RPTTravelChargeRankEntity> getServicePointIds(@Param("yearMonth") Integer yearMonth,
                                                        @Param("systemId")Integer systemId);

    Integer hasReportData(RPTTravelChargeRankSearchCondition condition);

    Page<RPTTravelChargeRankEntity> getTravelChargeRankList(RPTTravelChargeRankSearchCondition condition);


}
