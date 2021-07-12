package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.provider.rpt.entity.SysUserRegionEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface SysUserRegionMappingMapper {

    List<SysUserRegionEntity> getSysUserRegionMappingFromWebDB();

    void insert(SysUserRegionEntity entity);

    void deleteSysUserRegionMappings(@Param("quarter") String quarter,
                                     @Param("systemId") Integer systemId);


}
