package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTSMSQtyStatisticsEntity;
import com.kkl.kklplus.provider.rpt.entity.SmsQtyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface SMSQtyStatisticsMapper {

    List<RPTSMSQtyStatisticsEntity> getMessageType(@Param("systemId") Integer systemId,
                                                   @Param("startDate") Long startDate,
                                                   @Param("endDate") Long endDate);

    List<SmsQtyEntity> getSmsQtyRpt(@Param("systemId") Integer systemId,
                                    @Param("startDate") Long startDate,
                                    @Param("endDate") Long endDate);

    Integer hasReportData(@Param("systemId") Integer systemId,
                          @Param("startDate") Long startDate,
                          @Param("endDate") Long endDate);

    List<SmsQtyEntity> getSysSmsQtyData(@Param("startDate") Date startDate,
                                        @Param("endDate") Date endDate);

    void insert(SmsQtyEntity entity);

    void update(SmsQtyEntity entity);

    void delete(@Param("systemId") Integer systemId,
                @Param("beginDate") Long beginDate,
                @Param("endDate") Long endDate);
}
