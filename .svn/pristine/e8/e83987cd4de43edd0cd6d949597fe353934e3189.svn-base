package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerFinanceEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface CustomerFinanceRptMapper {

    List<RPTCustomerFinanceEntity> getCustomerFinance(@Param("customerIds") List<Long> customerIds);


    List<RPTCustomerFinanceEntity> getAllCustomerFinance();
}
