package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.CustomerRechargeSummaryRptService;
import com.kkl.kklplus.provider.rpt.service.DepositRechargeSummaryRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("depositRecharge")
public class DepositRechargeRptController {

    @Autowired
    private DepositRechargeSummaryRptService depositRechargeSummaryRptService;


    @ApiOperation("获取客戶质保金充值汇总")
    @PostMapping("getDepositRechargeSummary")
    public MSResponse<List<RPTCustomerRechargeSummaryEntity>> getDepositRechargeSummary(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTCustomerRechargeSummaryEntity> returnList = depositRechargeSummaryRptService.getDepositRechargeSummaryList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);
    }

    @ApiOperation("质保金充值明细")
    @PostMapping("getDepositRechargeDetails")
    public MSResponse<MSPage<RPTRechargeRecordEntity>> getDepositRechargeDetails(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        Page<RPTRechargeRecordEntity> servicePointList = depositRechargeSummaryRptService.getDepositRechargeDetailsByPage(search);
        MSPage<RPTRechargeRecordEntity> returnPage = new MSPage<>();
        if (servicePointList != null) {
            returnPage.setList(servicePointList);
            returnPage.setPageNo(servicePointList.getPageNum());
            returnPage.setPageSize(servicePointList.getPageSize());
            returnPage.setRowCount((int) servicePointList.getTotal());
            returnPage.setPageCount(servicePointList.getPages());
        }

        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
