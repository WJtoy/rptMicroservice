package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.KeFuOrderPlanDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("keFuOrderPlan")
public class KeFuOrderPlanDailyRptController {

    @Autowired
    private KeFuOrderPlanDailyRptService service;

    /**
     * 查询客服每日接单
     *
     * @return
     */
    @ApiOperation("获取客服每日接单数据")
    @PostMapping("keFuOrderPlanDaily")
    public MSResponse<List<RPTKeFuOrderPlanDailyEntity>> getAreaOrderPlanList(@RequestBody RPTKeFuOrderPlanDailySearch search) {
        List<RPTKeFuOrderPlanDailyEntity> objectMap =  service.getKeFuOrderAcceptDayRptData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, objectMap);
    }
}
