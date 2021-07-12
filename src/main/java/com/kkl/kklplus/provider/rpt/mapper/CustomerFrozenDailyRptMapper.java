package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerFrozenDailyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CustomerFrozenDailyRptMapper {

    /**
     * 查询系统中的客户每日冻结明细
     */
    Page<RPTCustomerFrozenDailyEntity> getCustomerFrozenDailyList(@Param("customerId") Long customerId,
                                                                    @Param("saleId") Long saleId,
                                                                    @Param("beginDate") Date beginDate,
                                                                    @Param("endDate") Date endDate,
                                                                    @Param("quarter") String quarter,
                                                                  @Param("subFlag")Integer subFlag);

    List<RPTCustomerFrozenDailyEntity> getCustomerFrozenDailyByList(@Param("customerId") Long customerId,
                                                                  @Param("saleId") Long saleId,
                                                                  @Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("quarter") String quarter,
                                                                    @Param("subFlag")Integer subFlag);





    List<RPTCustomerFrozenDailyEntity> getCustomerFrozenDailyIdsFromWebDB(@Param("currencyNos") List<String> currencyNos);


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
