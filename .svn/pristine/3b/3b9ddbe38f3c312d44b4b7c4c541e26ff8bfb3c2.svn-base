package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RptCustomerMonthOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.CustomerMonthPlanDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("customerMonthPlan")
public class CustomerMonthPlanDailyRptController {

    @Autowired
    CustomerMonthPlanDailyRptService customerMonthPlanDailyRptService;

    @ApiOperation("获取客户每月下单数据")
    @PostMapping("customerMonthPlanDaily")
    public MSResponse<List<RptCustomerMonthOrderEntity>> getCustomerOrderPlanList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RptCustomerMonthOrderEntity> list = customerMonthPlanDailyRptService.getCustomerMonthPlanDailyList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("获取客户每月下单图表数据")
    @PostMapping("getCustomerMonthPlanChartList")
    public MSResponse<Map<String, Object>> getCustomerOrderPlanChartList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        Map<String, Object> map =  customerMonthPlanDailyRptService.turnToCustomerOrderMonthPlanChart(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

}
