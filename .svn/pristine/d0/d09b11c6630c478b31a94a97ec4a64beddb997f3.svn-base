package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerReceivableSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface CustomerReceivableSummaryRptMapper {

    List<RPTCustomerSalesMappingEntity> getCustomerSalesList(@Param("systemId") Integer systemId);

    /**
     * 业务主管，取其下属的客户id
     * @param systemId
     * @param salesId
     * @return
     */
    List<RPTCustomerSalesMappingEntity> getCustomerSalesChargeList(@Param("systemId") Integer systemId,
                                                                   @Param("salesId") Long salesId,
                                                                   @Param("subFlag") Integer subFlag);


    List<RPTCustomerReceivableSummaryEntity> getCustomerReceivableByPage(@Param("yearMonth") Integer yearMonth,
                                                                         @Param("customerIds") List<Long> customerIds,
                                                                         @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                         @Param("systemId") Integer systemId);

    List<RPTCustomerReceivableSummaryEntity> getCustomerBalanceData(@Param("yearMonth") Integer yearMonth,
                                                                    @Param("customerIds") List<Long> customerIds,
                                                                    @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                    @Param("systemId") Integer systemId);

    List<RPTCustomerReceivableSummaryEntity>  getCurrentCreditList(@Param("customerIds") List<Long> customerIds);


    Integer hasReportData(@Param("yearMonth") Integer yearMonth,
                          @Param("customerId") Long customerId,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("systemId") Integer systemId);

    Integer hasReportList( @Param("customerId") Long customerId);
}
