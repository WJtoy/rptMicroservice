package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerNewOrderDailyRptEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;


@Mapper
public interface CustomerOrderDailyRptMapper {

    /**
     * 查询系统中的客户每日下单明细
     */
    Page<RPTCustomerNewOrderDailyRptEntity> getCustomerNewOrderDailyList(@Param("customerId") Long customerId,
                                                                         @Param("saleId") Long saleId,
                                                                         @Param("beginDate") Date beginDate,
                                                                         @Param("endDate") Date endDate,
                                                                         @Param("quarter") String quarter,
                                                                         @Param("subFlag")Integer subFlag);

    /**
     * 检查报表是否有数据
     */
    Integer hasReportData(@Param("customerId") Long customerId,
                          @Param("saleId") Long saleId,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("quarter") String quarter,
                          @Param("subFlag")Integer subFlag);

}
