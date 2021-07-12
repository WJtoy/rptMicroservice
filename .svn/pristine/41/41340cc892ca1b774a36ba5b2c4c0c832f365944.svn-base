package com.kkl.kklplus.provider.rpt.mapper;




import com.kkl.kklplus.entity.rpt.RPTCrushAreaEntity;
import com.kkl.kklplus.entity.rpt.RPTSpecialChargeAreaEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;
import java.util.Map;


@Mapper
public interface CrushAreaRptMapper {
    /**
     * 获取原表中的数据
     * @return
     */
    List<RPTCrushAreaEntity> getOldCrushData(@Param("startDate") Date startDate,
                                             @Param("endDate") Date endDate
    );

    List<RPTCrushAreaEntity> getCrushAreaList(@Param("yearMonth") Integer yearMonth,
                                              @Param("systemId") Integer systemId,
                                              @Param("productCategoryIds") List<Long> productCategoryIds,
                                              @Param("quarter") String quarter
    );

    void insertCrushRpt(RPTCrushAreaEntity entity);

    void updateCrushRpt(RPTCrushAreaEntity entity);

    void deleteCrushRpt(@Param("systemId") Integer systemId,
                                @Param("yearmonth") Integer yearmonth);

    List<Map<String,Long>> getCountyIds(@Param("yearMonth") Integer yearMonth,
                                        @Param("systemId") Integer systemId);

    Integer hasReportData(@Param("yearMonth") Integer yearMonth,
                          @Param("systemId") Integer systemId,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("quarter") String quarter
    );
}
