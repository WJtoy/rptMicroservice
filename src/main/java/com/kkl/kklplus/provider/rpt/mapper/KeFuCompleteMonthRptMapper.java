package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuCompletedMonthEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface KeFuCompleteMonthRptMapper {

    List<RPTKeFuCompletedMonthEntity> getKeFuCompleteMonthList(@Param("startYearMonth") Integer startYearMonth,
                                                               @Param("endYearMonth") Integer endYearMonth,
                                                               @Param("keFuId") Long keFuId,
                                                               @Param("productCategoryIds") List<Long> productCategoryIds,
                                                               @Param("systemId") Integer systemId);

    Integer hasReportData(@Param("startYearMonth") Integer startYearMonth,
                          @Param("endYearMonth") Integer endYearMonth,
                          @Param("keFuIds") List<Long> keFuIds,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("systemId") Integer systemId);
}
