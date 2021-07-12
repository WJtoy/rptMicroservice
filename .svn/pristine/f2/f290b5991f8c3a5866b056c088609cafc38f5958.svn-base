package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface CustomerSalesMappingMapper {

    List<RPTCustomerSalesMappingEntity> getCustomerSalesMappingFromWebDB();

    void insert(RPTCustomerSalesMappingEntity entity);

    void deleteCustomerSalesMappings(@Param("quarter") String quarter,
                                     @Param("systemId") Integer systemId);


}
