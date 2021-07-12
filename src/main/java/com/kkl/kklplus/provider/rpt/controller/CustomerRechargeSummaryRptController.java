package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.CustomerRechargeSummaryRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("customerRechargeSummary")
public class CustomerRechargeSummaryRptController {

    @Autowired
    private CustomerRechargeSummaryRptService customerRechargeSummaryRptService;


    @ApiOperation("获取客戶充值汇总")
    @PostMapping("getCustomerRechargeSummary")
    public MSResponse<List<RPTCustomerRechargeSummaryEntity>> getCustomerRechargeSummary(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTCustomerRechargeSummaryEntity> returnList = customerRechargeSummaryRptService.getCustomerRechargeSummaryList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);
    }
}
