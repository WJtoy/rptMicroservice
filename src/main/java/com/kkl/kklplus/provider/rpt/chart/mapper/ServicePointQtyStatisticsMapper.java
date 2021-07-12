package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.chart.entity.RPTServicePointQtyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface ServicePointQtyStatisticsMapper {

    RPTServicePointQtyEntity getServicePointQty(@Param("systemId") Integer systemId,
                                                @Param("beginDate") Long beginDate,
                                                @Param("endDate") Long endDate,
                                                @Param("productCategoryId") Long productCategoryId);

    List<RPTServicePointQtyEntity> getServicePointProductCategoryQty(@Param("systemId") Integer systemId,
                                                                     @Param("beginDate") Long beginDate,
                                                                     @Param("endDate") Long endDate,
                                                                     @Param("productCategoryIds") List<Long> productCategoryIds);

    List<RPTServicePointQtyEntity> getServicePointAutoPlanQty(@Param("systemId") Integer systemId,
                                                              @Param("beginDate") Long beginDate,
                                                              @Param("endDate") Long endDate,
                                                              @Param("productCategoryIds") List<Long> productCategoryIds);

    void insertServicePointQtyData(RPTServicePointQtyEntity entity);

    void deleteServicePointQtyFromRptDB(@Param("systemId") Integer systemId,
                                        @Param("startDate") Long startDate,
                                        @Param("endDate") Long endDate);

}
