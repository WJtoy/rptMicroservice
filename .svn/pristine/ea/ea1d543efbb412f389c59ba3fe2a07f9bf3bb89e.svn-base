package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTOrderDailyWorkEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.service.RPTOrderDailyWorkService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("rptCreatedOrder")
public class RptCreatedOrderController {
    @Autowired
    RPTOrderDailyWorkService rptOrderDailyWorkService;


    @ApiOperation("每日工单统计")
    @PostMapping("getCreatedOrderList")
    public MSResponse<List<RPTOrderDailyWorkEntity>> getCreatedOrderList(@RequestBody RPTCompletedOrderDetailsSearch search){
        List<RPTOrderDailyWorkEntity> returnList = rptOrderDailyWorkService.getCreatedOrderList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);


    }
}
