package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTComplainStatisticsDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.provider.rpt.service.ComplainStatisticsDailyrPptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("complainStatistics")
public class ComplainStatisticsDailyRptController {

    @Autowired
    private ComplainStatisticsDailyrPptService complainStatisticsDailyrPptService;

    @ApiOperation("获取每日投诉数据")
    @PostMapping("getComplainStatisticsDailyList")
    public MSResponse<List<RPTComplainStatisticsDailyEntity>> getComplainStatisticsDailyList(@RequestBody RPTComplainStatisticsDailySearch search) {
        List<RPTComplainStatisticsDailyEntity> list = complainStatisticsDailyrPptService.getComplainStatisticsDailyList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

    @ApiOperation("获取每日投诉图表数据")
    @PostMapping("getComplainStatisticsDailyChart")
    public MSResponse<Map<String, Object>> getComplainStatisticsDailyChart(@RequestBody RPTComplainStatisticsDailySearch search) {
        Map<String, Object> map =  complainStatisticsDailyrPptService.getComplainStatisticsDailyChartList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
}
