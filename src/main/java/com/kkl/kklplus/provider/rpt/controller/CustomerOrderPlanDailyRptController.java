package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.CustomerOrderPlanDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("customerOrderPlan")
public class CustomerOrderPlanDailyRptController {
    @Autowired
    private CustomerOrderPlanDailyRptService customerOrderPlanDailyRptService;

    @ApiOperation("获取客户每日下单数据")
    @PostMapping("customerOrderPlanDaily")
    public MSResponse<List<RPTCustomerOrderPlanDailyEntity>> getCustomerOrderPlanList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTCustomerOrderPlanDailyEntity> list =  customerOrderPlanDailyRptService.getCustomerOrderPlanDailyRptData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("获取客户每日下单图表数据")
    @PostMapping("customerOrderPlanDailyChart")
    public MSResponse<Map<String, Object>> getCustomerOrderPlanChartList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        Map<String, Object> map =  customerOrderPlanDailyRptService.turnToCustomerOrderPlanDailyChartInformation(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

}
