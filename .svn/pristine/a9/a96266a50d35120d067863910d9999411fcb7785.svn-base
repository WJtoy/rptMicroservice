package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTAreaOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTAreaOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.AreaOrderPlanDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("areaOrderPlan")
public class AreaOrderPlanDailyRptController {

    @Autowired
    private AreaOrderPlanDailyRptService areaOrderPlanDailyRptService;

    /**
     * 查询省市区每日下单
     *
     * @return
     */
    @ApiOperation("获取省市区每日下单数据")
    @PostMapping("areaOrderPlanDaily")
    public MSResponse<Map<String,List<RPTAreaOrderPlanDailyEntity>>> getAreaOrderPlanList(@RequestBody RPTAreaOrderPlanDailySearch search) {
        Map<String,List<RPTAreaOrderPlanDailyEntity>> objectMap =  areaOrderPlanDailyRptService.getAreaOrderPlanDailyRptData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, objectMap);
    }
}
