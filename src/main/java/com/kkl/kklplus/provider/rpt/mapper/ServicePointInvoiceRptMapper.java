package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface ServicePointInvoiceRptMapper {

    Page<RPTServicePointInvoiceEntity> getServicePointInvoicePage(@Param("servicePointId") Long servicePointId,
                                                                  @Param("withdrawNo") String withdrawNo,
                                                                  @Param("paymentType") Integer paymentType,
                                                                  @Param("bank") Integer bank,
                                                                  @Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("beginInvoiceDate") Date beginInvoiceDate,
                                                                  @Param("endInvoiceDate") Date endInvoiceDate,
                                                                  @Param("quarters")  List<String> quarters,
                                                                  @Param("status")  Integer status,
                                                                  @Param("pageNo")  Integer pageNo,
                                                                  @Param("pageSize")  Integer pageSize);


    List<RPTServicePointInvoiceEntity> getServicePointInvoiceList(@Param("servicePointId") Long servicePointId,
                                                                  @Param("withdrawNo") String withdrawNo,
                                                                  @Param("paymentType") Integer paymentType,
                                                                  @Param("bank") Integer bank,
                                                                  @Param("beginDate") Date beginDate,
                                                                  @Param("endDate") Date endDate,
                                                                  @Param("beginInvoiceDate") Date beginInvoiceDate,
                                                                  @Param("endInvoiceDate") Date endInvoiceDate,
                                                                  @Param("quarters")  List<String> quarters,
                                                                  @Param("status")  Integer status);

    Integer hasReportData(@Param("servicePointId") Long servicePointId,
                          @Param("withdrawNo") String withdrawNo,
                          @Param("paymentType") Integer paymentType,
                          @Param("bank") Integer bank,
                          @Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("beginInvoiceDate") Date beginInvoiceDate,
                          @Param("endInvoiceDate") Date endInvoiceDate,
                          @Param("quarters")  List<String> quarters,
                          @Param("status")  Integer status);
}
