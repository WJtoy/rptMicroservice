package com.kkl.kklplus.provider.rpt.chart.mapper;

import com.kkl.kklplus.provider.rpt.chart.entity.RPTServicePointStreetQtyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface ServicePointStreetQtyMapper {


    RPTServicePointStreetQtyEntity getServicePointStreetQty(@Param("systemId") Integer systemId,
                                                            @Param("beginDate") Long beginDate,
                                                            @Param("endDate") Long endDate,
                                                            @Param("productCategoryId") Long productCategoryId);

    List<RPTServicePointStreetQtyEntity> getServicePointProductCategoryQty(@Param("systemId") Integer systemId,
                                                                           @Param("beginDate") Long beginDate,
                                                                           @Param("endDate") Long endDate,
                                                                           @Param("productCategoryIds") List<Long> productCategoryIds);

    RPTServicePointStreetQtyEntity getServicePointStreetAutoPlanQty(@Param("systemId") Integer systemId,
                                                                    @Param("beginDate") Long beginDate,
                                                                    @Param("endDate") Long endDate,
                                                                    @Param("productCategoryId") Long productCategoryId);


    void insertServicePointStreetQtyData(RPTServicePointStreetQtyEntity entity);


    void deleteServicePointStreetQtyFromRptDB(@Param("systemId") Integer systemId,
                                              @Param("startDate") Long startDate,
                                              @Param("endDate") Long endDate);
}
