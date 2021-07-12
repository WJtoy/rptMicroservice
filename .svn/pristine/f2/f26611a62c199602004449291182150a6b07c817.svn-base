package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;


import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteEntity;
import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteSearch;
import com.kkl.kklplus.provider.rpt.service.EveryDayCompleteService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

/**
 * @Auther wj
 * @Date 2021/5/21 16:50
 * 每日完工时效报表
 */
@RestController
@RequestMapping("everyDayComplete")
public class EveryDayCompleteController {

    @Autowired
    private EveryDayCompleteService everyDayCompleteService;


    @ApiOperation("每日完成时效")
    @PostMapping("everyDayCompleteRate")
    public MSResponse<Map<String,List<RPTEveryDayCompleteEntity>>> getCustomerOrderPlanList(@RequestBody RPTEveryDayCompleteSearch search) {
        Map<String,List<RPTEveryDayCompleteEntity>> list = everyDayCompleteService.getEveryDayComplete(search);
        return new MSResponse<>(MSErrorCode.SUCCESS,list);
    }



}
