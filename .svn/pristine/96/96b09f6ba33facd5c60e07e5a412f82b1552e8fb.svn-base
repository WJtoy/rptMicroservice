package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderTimeSearch;
import com.kkl.kklplus.provider.rpt.entity.ChargeBaseEntity;
import com.kkl.kklplus.provider.rpt.entity.CustomerOrderTimeRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CustomerOrderTimeRptMapper {

    List<ChargeBaseEntity> getRecordCount(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getLessSix(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getMoreSix(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getClose12to24(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getCloseLessTwelve(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getClose24to48(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getClose48to72(RPTCustomerOrderTimeSearch search);

    List<ChargeBaseEntity> getCloseMore72(RPTCustomerOrderTimeSearch search);


    List<CustomerOrderTimeRptEntity> getCustomerOrderTimeData(@Param("beginDate") Date beginDate,
                                                              @Param("endDate") Date endDate);

    void insertCustomerOrderTimeData(@Param("list") List<CustomerOrderTimeRptEntity> entity);


    void deleteCustomerOrderTimeData(@Param("systemId") Integer systemId,
                                     @Param("beginDate") Long beginDate,
                                     @Param("endDate") Long endDate,
                                     @Param("quarter") String quarter);

    List<LongTwoTuple> getCompletedOrderIds(@Param("quarter") String quarter,
                                            @Param("systemId") Integer systemId,
                                            @Param("beginChargeDate") Long beginDate,
                                            @Param("endChargeDate") Long endDate);

    Integer hasReportData(RPTCustomerOrderTimeSearch search);

}
