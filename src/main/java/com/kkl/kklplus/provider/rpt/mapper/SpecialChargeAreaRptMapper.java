package com.kkl.kklplus.provider.rpt.mapper;




import com.kkl.kklplus.entity.rpt.RPTSpecialChargeAreaEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;
import java.util.Map;


@Mapper
public interface SpecialChargeAreaRptMapper {
    /**
     * 获取原表中的数据
     * @return
     */
    List<RPTSpecialChargeAreaEntity> getOldSpecialData(@Param("startDate")Date startDate,
                                                       @Param("endDate")Date endDate
                                                         );

    List<RPTSpecialChargeAreaEntity> getSpecialChargeAreaList(@Param("yearMonth")Integer yearMonth,
                                                              @Param("areaId") Long areaId,
                                                              @Param("systemId") Integer systemId,
                                                              @Param("productCategoryIds")List<Long> productCategoryIds
                                                              );

    void insertSpecialChargeRpt(RPTSpecialChargeAreaEntity entity);

    void updateSpecialChargeRpt(RPTSpecialChargeAreaEntity entity);

    void deleteSpecialChargeRpt(@Param("systemId") Integer systemId,
                            @Param("yearmonth") Integer yearmonth);

    List<Map<String,Long>> getCountyIds(@Param("yearMonth") Integer yearMonth,
                                        @Param("systemId")Integer systemId);

    Integer hasReportData(@Param("yearMonth") Integer yearMonth,
                          @Param("areaId") Long areaId,
                          @Param("systemId") Integer systemId,
                          @Param("productCategoryIds") List<Long> productCategoryIds
                         );
}
