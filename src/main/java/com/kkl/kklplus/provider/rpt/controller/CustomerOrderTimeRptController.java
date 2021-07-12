package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderTimeSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerOrderTimeRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("customerOrderTime")
public class CustomerOrderTimeRptController {
    @Autowired
    private CustomerOrderTimeRptService customerOrderTimeRptService;
    /**
     * 条件查询
     */
    @ApiOperation("获取客户工单时效")
    @PostMapping("getCustomerOrderTimeRptList")
    public MSResponse<List<RPTCustomerOrderTimeEntity>> getCustomerOrderTimeRptList(@RequestBody RPTCustomerOrderTimeSearch reminderSearch) {
        List<RPTCustomerOrderTimeEntity> reminderList = customerOrderTimeRptService.getCustomerOrderTimeRptData(reminderSearch);
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }

    @ApiOperation("获取客户工单时效图表数据")
    @PostMapping("getCustomerOrderTimeChartList")
    public MSResponse<Map<String, Object>> getCustomerOrderTimeChartList(@RequestBody RPTCustomerOrderTimeSearch search) {
        Map<String, Object> map =  customerOrderTimeRptService.getCustomerOrderTimeChart(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
}
