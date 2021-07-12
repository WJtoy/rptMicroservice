package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCustomerRevenueEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerRevenueSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerRevenueRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("customerRevenue")
public class CustomerRevenueRptController {
    @Autowired
    private CustomerRevenueRptService customerRevenueRptService;
    /**
     * 条件查询
     */
    @ApiOperation("获取客户营收统计")
    @PostMapping("getCustomerRevenueRpt")
    public MSResponse<List<RPTCustomerRevenueEntity>> getCustomerRevenueRptList(@RequestBody RPTCustomerRevenueSearch reminderSearch) {
        List<RPTCustomerRevenueEntity> reminderList = customerRevenueRptService.getCustomerRevenueList(reminderSearch);
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }

    @ApiOperation("获取客户营收排名")
    @PostMapping("getCustomerRevenueChartList")
    public MSResponse<Map<String, Object>> getCustomerRevenueChartList(@RequestBody RPTCustomerRevenueSearch search) {
        Map<String, Object> map =  customerRevenueRptService.turnToChartInformation(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
}
