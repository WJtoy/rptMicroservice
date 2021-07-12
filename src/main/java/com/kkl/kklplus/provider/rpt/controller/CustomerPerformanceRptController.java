package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTSalesPerfomanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.CustomerPerformanceRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("salesPerformanceReport")
public class CustomerPerformanceRptController {


    @Autowired
    CustomerPerformanceRptService customerPerformanceRptService;

    @ApiOperation("获取业务员业绩")
    @PostMapping("getSalesPerformanceList")
    public MSResponse<List<RPTSalesPerfomanceEntity>> getSalesPerformanceList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTSalesPerfomanceEntity> list = customerPerformanceRptService.getSalesPerformanceByList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }


    @ApiOperation("获取业务员业绩明細")
    @PostMapping("getCustomerPerformanceList")
    public MSResponse<List<RPTSalesPerfomanceEntity>> getCustomerPerformanceList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTSalesPerfomanceEntity> list = customerPerformanceRptService.getCustomerPerformanceList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

}
