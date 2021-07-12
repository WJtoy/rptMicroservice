package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.provider.rpt.entity.SysUserRegionEntity;
import com.kkl.kklplus.provider.rpt.entity.UserCustomerEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface SysUserCustomerMappingMapper {

    List<UserCustomerEntity> getSysUserCustomerMappingFromWebDB();

    void insert(UserCustomerEntity entity);

    void deleteSysUserCustomerMappings(@Param("quarter") String quarter,
                                       @Param("systemId") Integer systemId);


}
