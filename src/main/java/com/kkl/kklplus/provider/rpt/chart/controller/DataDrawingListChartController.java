package com.kkl.kklplus.provider.rpt.chart.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainChartEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.service.*;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;


@RestController
@RequestMapping("dataDrawingList")
public class DataDrawingListChartController {

    @Autowired
    private KeFuCompleteTimeChartService keFuCompleteTimeChartService;

    @Autowired
    private CustomerPlanChartService orderDataChartService;

    @Autowired
    private CustomerReminderChartService customerReminderChartService;

    @Autowired
    private CustomerComplainChartService customerComplainChartService;

    @Autowired
    private ServicePointQtyStatisticsService servicePointQtyStatisticsService;

    @Autowired
    private ServicePointStreetQtyService servicePointStreetQtyService;

    @Autowired
    private IncurExpenseChartService incurExpenseChartService;

    @Autowired
    private OrderQtyDailyChartService orderSituationService;

    @Autowired
    private OrderCrushQtyChartService orderCrushQtyChartService;

    @Autowired
    private OrderPlanDailyChartService orderPlanDailyChartService;

    @ApiOperation("获取客服完工时效图表数据")
    @PostMapping("getKeFuCompleteTimeInstallChartList")
    public MSResponse<Map<String, Object>> getKeFuCompleteTimeInstallChartList(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = keFuCompleteTimeChartService.getKeFuCompleteTimeRptData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

    @ApiOperation("获取下单图表数据")
    @PostMapping("getOrderDataChartList")
    public MSResponse<Map<String, Object>> getOrderDataChartList(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = orderDataChartService.getOrderDataChartData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
    @ApiOperation("获取工单图表数据")
    @PostMapping("getOrderQtyDailyChartData")
    public MSResponse<Map<String, Object>> getOrderQtyDailyChartData(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = orderSituationService.getOrderQtyDailyChartData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
    @ApiOperation("获取每日催单图表数据")
    @PostMapping("getCustomerReminderChart")
    public MSResponse<RPTCustomerReminderEntity> getCustomerReminderChart(@RequestBody RPTDataDrawingListSearch search) {
        RPTCustomerReminderEntity entity = customerReminderChartService.getCustomerReminderChartData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, entity);
    }

    @ApiOperation("获取客户投诉图表数据")
    @PostMapping("getCustomerComplainChart")
    public MSResponse<RPTCustomerComplainChartEntity> getCustomerComplainChart(@RequestBody RPTDataDrawingListSearch search) {
        RPTCustomerComplainChartEntity entity = customerComplainChartService.getCustomerComplain(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, entity);
    }

    @ApiOperation("获取网点数量图表数据")
    @PostMapping("getServicePointQtyChart")
    public MSResponse<Map<String, Object>> getServicePointQtyChart(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = servicePointQtyStatisticsService.getServicePointQtyData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

    @ApiOperation("获取网点街道数量图表数据")
    @PostMapping("getServicePointStreetQtyChart")
    public MSResponse<Map<String, Object>> getServicePointStreetQtyChart(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = servicePointStreetQtyService.getServicePointStreetQtyData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
    @ApiOperation("获取支出费用图表数据")
    @PostMapping("getIncurExpenseChart")
    public MSResponse<List<Double>> getIncurExpenseChart(@RequestBody RPTDataDrawingListSearch search) {
        List<Double> list = incurExpenseChartService.getIncurExpenseChart(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("获取突击单数量图表数据")
    @PostMapping("getOrderCrushQtyChart")
    public MSResponse<Map<String, Object>> getOrderCrushQtyChart(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = orderCrushQtyChartService.getOrderCrushQtyData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

    @ApiOperation("获取日下单明细图表数据")
    @PostMapping("getOrderPlanDailyChart")
    public MSResponse<Map<String, Object>> getOrderPlanDailyChart(@RequestBody RPTDataDrawingListSearch search) {
        Map<String, Object> map = orderPlanDailyChartService.getOrderPlanDailyChartData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

}
