package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointInvoiceSearch;
import com.kkl.kklplus.provider.rpt.service.ServicePointPaymentSummaryRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("servicePointPayment")
public class ServicePointPaymentSummaryRptController {

    @Autowired
    private ServicePointPaymentSummaryRptService servicePointPaymentSummaryRptService;


    @ApiOperation("获取网点付款汇总数据")
    @PostMapping("servicePointPaymentSummary")
    public MSResponse<List<RPTServicePointInvoiceEntity>> getServicePointPaymentSummary(@RequestBody RPTServicePointInvoiceSearch search) {
        List<RPTServicePointInvoiceEntity> returnList = servicePointPaymentSummaryRptService.getServicePointPaymentSummaryList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);
    }
}
