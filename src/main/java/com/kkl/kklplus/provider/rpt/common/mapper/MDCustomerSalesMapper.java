package com.kkl.kklplus.provider.rpt.common.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface MDCustomerSalesMapper {

    List<Long> getSalesIdList(@Param("systemId") Integer systemId);

}
