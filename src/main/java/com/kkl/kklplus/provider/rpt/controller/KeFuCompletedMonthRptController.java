package com.kkl.kklplus.provider.rpt.controller;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuCompletedMonthEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.provider.rpt.service.KeFuCompletedMonthRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("KeFuCompletedMonth")
public class KeFuCompletedMonthRptController {

    @Autowired
    private KeFuCompletedMonthRptService keFuCompletedMonthRptService;

    @ApiOperation("客服每月完工单")
    @PostMapping("getKeFuCompletedMonthInfo")
    public MSResponse<List<RPTKeFuCompletedMonthEntity>> getKeFuCompletedMonthInfo(@RequestBody RPTGradedOrderSearch search) {
        List<RPTKeFuCompletedMonthEntity> list = keFuCompletedMonthRptService.getKeFuCompletedMonthList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("客服每月完工单图表数据")
    @PostMapping("getKeFuCompletedMonthChartList")
    public MSResponse<Map<String, Object>> getKeFuCompletedMonthChartList(@RequestBody RPTGradedOrderSearch search) {
        Map<String, Object> map =  keFuCompletedMonthRptService.turnToKeFuCompletedMonthPlanChart(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

}
