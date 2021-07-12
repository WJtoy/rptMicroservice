package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerReceivableSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointInvoiceSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerReceivableSummaryRptService;
import com.kkl.kklplus.provider.rpt.service.ServicePointInvoiceRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("customerReceivableSummary")
public class CustomerReceivableSummaryRptController {

    @Autowired
    private CustomerReceivableSummaryRptService customerReceivableSummaryRptService;

    @ApiOperation("获取网点付款清单")
    @PostMapping("getCustomerReceivableSummaryByPage")
    public MSResponse<MSPage<RPTCustomerReceivableSummaryEntity>> getCustomerReceivableSummaryByPage(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        MSPage<RPTCustomerReceivableSummaryEntity> returnPage = customerReceivableSummaryRptService.getCustomerReceivableSummaryMonth(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
