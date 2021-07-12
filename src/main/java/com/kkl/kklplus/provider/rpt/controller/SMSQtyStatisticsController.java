package com.kkl.kklplus.provider.rpt.controller;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTSMSQtyStatisticsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTSMSQtyStatisticsSearch;
import com.kkl.kklplus.provider.rpt.service.SMSQtyStatisticsService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("smsQtyStatistics")
public class SMSQtyStatisticsController {

    @Autowired
    private SMSQtyStatisticsService SMSQtyStatisticsService;
    /**
     * 条件查询
     */
    @ApiOperation("获取短信数量")
    @PostMapping("getSMSQtyStatisticsRptList")
    public MSResponse<List<RPTSMSQtyStatisticsEntity>> getSMSQtyStatisticsRptList(@RequestBody RPTSMSQtyStatisticsSearch reminderSearch) {
        List<RPTSMSQtyStatisticsEntity> reminderList = SMSQtyStatisticsService.getMessageNumberRpt(reminderSearch);
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }

    @ApiOperation("获取短信数量图表数据")
    @PostMapping("getSMSQtyStatisticsChartList")
    public MSResponse<Map<String, Object>> getSMSQtyStatisticsRptChartList(@RequestBody RPTSMSQtyStatisticsSearch search) {
        Map<String, Object> map =  SMSQtyStatisticsService.turnToChartInformation(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, map);
    }
}
