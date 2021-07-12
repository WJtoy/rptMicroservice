package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTDispatchOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.service.DispatchListInformationRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("dispatchList")
public class DispatchListInformationRptController {

    @Autowired
    DispatchListInformationRptService dispatchListInformationRptService;

    @ApiOperation("接派单来源统计")
    @PostMapping("dispatchListInformation")
    public MSResponse<List<RPTDispatchOrderEntity>> getCustomerOrderPlanList(@RequestBody RPTCompletedOrderDetailsSearch search) {
        List<RPTDispatchOrderEntity> list = dispatchListInformationRptService.getDispatchListInformation(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("获取接派单图表数据")
    @PostMapping("getDispatchListInforChart")
    public MSResponse<Map<String, Object>> getDispatchListInforChart(@RequestBody RPTCompletedOrderDetailsSearch search) {
        Map<String, Object> map =  dispatchListInformationRptService.getDispatchListInformationChart(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }

}
