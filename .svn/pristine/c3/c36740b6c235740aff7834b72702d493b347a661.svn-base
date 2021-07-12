package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTGradeQtyDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.entity.CustomerPerformanceRptEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface GradeQtyDailyRptMapper {

    //按天统计客评类型数量
    List<GradeQtyEntity> getGradeQtyList(@Param("startDate") Date startDate,
                                         @Param("endDate") Date endDate);

    void insertGradeQty(GradeQtyRptEntity entity);


    void delete(@Param("beginDate") Long beginDate,
                @Param("endDate") Long endDate,
                @Param("systemId") Integer systemId);

    List<RPTGradeQtyDailyEntity> getGradeQtyDailyList(RPTCustomerOrderPlanDailySearch search);

    Integer hasReportData(RPTCustomerOrderPlanDailySearch search);

    List<LongThreeTuple> getGradeQtyDailyIds(@Param("systemId") Integer systemId,
                                             @Param("startDate") long startDate,
                                             @Param("endDate") long endDate);

    void updateGradeQty(GradeQtyRptEntity entity);
}
