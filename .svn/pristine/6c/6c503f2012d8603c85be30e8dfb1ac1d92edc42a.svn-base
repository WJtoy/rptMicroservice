package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTChargeDailyEntity;
import com.kkl.kklplus.provider.rpt.entity.ChargeDailyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ChargeDailyRptMapper {

    /**
     * 获取每日完工的数量
     */
    List<RPTAreaCompletedDailyEntity> getCompletedDaily(@Param("startYearMonth") Integer startYearMonth,
                                                        @Param("systemId") Integer systemId
    );

    /**
     * 获取手动对账的数量
     *
     * @param
     * @return
     */
    List<ChargeDailyEntity> getManualChargeDaily(@Param("startDate") Date startDate,
                                                 @Param("lastDate") Date lastDate
    );


    /**
     * 获取自动对账的数量
     */
    List<ChargeDailyEntity> getAutoChargeDaily(@Param("startDate") Date startDate,
                                              @Param("lastDate") Date lastDate);



    Integer getCompletedSum(@Param("startYearMonth") Integer startYearMonth,
                            @Param("systemId") Integer systemId);



    Integer getManualChargeSum(@Param("startDate") Date startDate,
                               @Param("lastDate") Date lastDate);


    Integer getAutoChargeSum(@Param("startDate") Date startDate,
                             @Param("lastDate") Date lastDate);
}
