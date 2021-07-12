package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderCancelledDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderCancelledDailySearch;
import com.kkl.kklplus.provider.rpt.service.KeFuOrderCancelledDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("keFuOrderCancelled")
public class KeFuOrderCancelledDailyRptController {
    @Autowired
    private KeFuOrderCancelledDailyRptService service;

    /**
     * 查询客服每日退单
     *
     * @return
     */
    @ApiOperation("获取客服每日退单数据")
    @PostMapping("keFuOrderCancelledDaily")
    public MSResponse<List<RPTKeFuOrderCancelledDailyEntity>> getKeFuOrderCancelledList(@RequestBody RPTKeFuOrderCancelledDailySearch search) {
        List<RPTKeFuOrderCancelledDailyEntity> list =  service.getKeFuOrderCancelledDailyRptData(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }
}
