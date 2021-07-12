package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ServicePointPaymentSummaryRptMapper {

    List<RPTServicePointInvoiceEntity> getServicePointPaymentList(@Param("paymentType") Integer paymentType,
                                                                  @Param("bank") Integer bank,
                                                                  @Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("quarter") String quarter);

    Integer hasReportData(@Param("paymentType") Integer paymentType,
                          @Param("bank") Integer bank,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("quarter") String quarter);
}
