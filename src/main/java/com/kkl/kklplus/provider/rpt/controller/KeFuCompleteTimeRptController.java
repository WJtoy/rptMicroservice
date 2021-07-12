package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.provider.rpt.service.KeFuCompleteTimeRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("keFuCompleteTime")
public class KeFuCompleteTimeRptController {

    @Autowired
    private KeFuCompleteTimeRptService keFuCompleteTimeRptService;
    /**
     * 条件查询
     */
    @ApiOperation("获取客服完工时效")
    @PostMapping("getKeFuCompleteTimeRptList")
    public MSResponse<List<RPTKeFuCompleteTimeEntity>> getKeFuCompleteTimeRptList(@RequestBody RPTKeFuCompleteTimeSearch reminderSearch) {
        List<RPTKeFuCompleteTimeEntity> reminderList = keFuCompleteTimeRptService.getKeFuCompleteTimeRptData(reminderSearch);
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }

    @ApiOperation("获取客服完工时效图表数据")
    @PostMapping("getKeFuCompleteTimeChartList")
    public MSResponse<Map<String, Object>> getKeFuCompleteTimeRptChartList(@RequestBody RPTKeFuCompleteTimeSearch search) {
        Map<String, Object> map =  keFuCompleteTimeRptService.turnToChartInformationNew(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
}
