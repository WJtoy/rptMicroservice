package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.RechargeRecordRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("getRechargeRecord")
public class RechargeRecordRptController {

    @Autowired
    private RechargeRecordRptService rechargeRecordRptService;

    @ApiOperation("充值明细")
    @PostMapping("getRechargeRecordByPage")
    public MSResponse<MSPage<RPTRechargeRecordEntity>> getRechargeRecordByPage(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        Page<RPTRechargeRecordEntity> servicePointList = rechargeRecordRptService.getRechargeRecordByPage(search);
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
