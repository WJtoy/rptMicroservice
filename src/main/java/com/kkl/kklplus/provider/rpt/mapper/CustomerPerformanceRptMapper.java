package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import com.kkl.kklplus.entity.rpt.RPTSalesPerfomanceEntity;
import com.kkl.kklplus.provider.rpt.entity.CustomerPerformanceRptEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;
import io.swagger.models.auth.In;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;
import java.util.Map;

@Mapper
public interface CustomerPerformanceRptMapper {

    List<CustomerPerformanceRptEntity> getFinishQtyList(@Param("beginCreatDate") long beginCreatDate,
                                                        @Param("endCreatDate") long endCreatDate,
                                                        @Param("systemId") Integer systemId,
                                                        @Param("quarter") String quarter);

    List<CustomerPerformanceRptEntity> getReturnQtyList(@Param("beginCreatDate") long beginCreatDate,
                                                        @Param("endCreatDate") long endCreatDate,
                                                        @Param("systemId") Integer systemId,
                                                        @Param("quarter") String quarter);

    List<CustomerPerformanceRptEntity> getCancelQtyList(@Param("beginCreatDate") long beginCreatDate,
                                                        @Param("endCreatDate") long endCreatDate,
                                                        @Param("systemId") Integer systemId,
                                                        @Param("quarter") String quarter);

    List<CustomerPerformanceRptEntity> getCreateQtyList(@Param("beginCreatDate") long beginCreatDate,
                                                        @Param("endCreatDate") long endCreatDate,
                                                        @Param("systemId") Integer systemId,
                                                        @Param("quarter") String quarter);

    List<LongThreeTuple> getCustomerOrderQtyMonthlyIds(@Param("systemId") Integer systemId,
                                                       @Param("yearmonth") Integer yearmonth);

    void insertCustomerPerformance(CustomerPerformanceRptEntity entity);

    void delete(@Param("yearMonth") Integer yearMonth,
                @Param("systemId") Integer systemId);

    List<RPTCustomerSalesMappingEntity> getCustomerSalesList(@Param("systemId") Integer systemId);

    /**
     * 如是业务主管，取其业务下属的业务员id和客户id
     * @param systemId
     * @param salesId
     * @return
     */
    List<RPTCustomerSalesMappingEntity> getCustomerSalesChargeList(@Param("systemId") Integer systemId,@Param("salesId") Long salesId);

    List<RPTSalesPerfomanceEntity> getCustomerPerformanceByList(@Param("yearMonth") Integer yearMonth,
                                                                @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                @Param("systemId") Integer systemId,
                                                                @Param("salesId") Long salesId,
                                                                @Param("subFlag")Integer subFlag);

    List<RPTSalesPerfomanceEntity> getSalesManAchievementDataNew(@Param("yearMonth") Integer yearMonth,
                                                                 @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                 @Param("systemId") Integer systemId,
                                                                 @Param("salesId") Long salesId,
                                                                 @Param("subFlag") Integer subFlag);

    List<RPTSalesPerfomanceEntity> getSalesNoFinishOrderQty(@Param("yearMonth") Integer yearMonth,
                                                            @Param("productCategoryIds") List<Long> productCategoryIds,
                                                            @Param("systemId") Integer systemId,
                                                            @Param("salesId") Long salesId,
                                                            @Param("subFlag")Integer subFlag);

    List<RPTSalesPerfomanceEntity> getCustomerNoFinishOrderQtyNew(@Param("yearmonth") Integer yearMonth,
                                                                  @Param("salesId") Long salesId,
                                                                  @Param("productCategoryIds") List<Long> productCategoryIds,
                                                                  @Param("systemId") Integer systemId,
                                                                  @Param("subFlag")Integer subFlag);

    Integer hasReportData(@Param("yearMonth") Integer yearMonth,
                          @Param("productCategoryIds") List<Long> productCategoryIds,
                          @Param("systemId") Integer systemId,
                          @Param("salesId") Long salesId,
                          @Param("subFlag")Integer subFlag);

    void updateCustomerOrderQtyMonthly(CustomerPerformanceRptEntity entity);

}
