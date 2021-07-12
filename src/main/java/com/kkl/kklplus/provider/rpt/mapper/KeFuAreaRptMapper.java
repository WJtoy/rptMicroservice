package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTKeFuAreaEntity;
import com.kkl.kklplus.provider.rpt.entity.KeFuAreaEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface KeFuAreaRptMapper {

    List<RPTKeFuAreaEntity> getKeFuAreaList( @Param("systemId") Integer systemId);

    List<KeFuAreaEntity> getKeFuProvinceList(@Param("systemId") Integer systemId);

    List<KeFuAreaEntity> getKeFuCityList( @Param("systemId") Integer systemId);

    List<KeFuAreaEntity> getKeFuAreasList( @Param("systemId") Integer systemId);

    List<KeFuAreaEntity> getKeFuCountryList( @Param("systemId") Integer systemId);

}
