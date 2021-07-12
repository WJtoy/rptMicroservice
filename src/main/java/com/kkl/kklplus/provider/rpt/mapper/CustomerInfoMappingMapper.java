package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import com.kkl.kklplus.provider.rpt.entity.CustomerInfoEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface CustomerInfoMappingMapper {

    List<CustomerInfoEntity> getCustomerInfoMappingFromWebDB();

    void insert(CustomerInfoEntity entity);

    void deleteCustomerInfoMappings(@Param("quarter") String quarter,
                                     @Param("systemId") Integer systemId);


}
