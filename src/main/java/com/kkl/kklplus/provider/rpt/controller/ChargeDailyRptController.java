package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTChargeDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTComplainStatisticsDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.provider.rpt.service.ChargeDailyRptService;
import com.kkl.kklplus.provider.rpt.service.ComplainStatisticsDailyrPptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.lang.reflect.InvocationTargetException;
import java.util.List;

@RestController
@RequestMapping("chargeDailyList")
public class ChargeDailyRptController {

    @Autowired
    private ChargeDailyRptService chargeDailyRptService;

    @ApiOperation("每日对账统计")
    @PostMapping("getChargeDailyList")
    public MSResponse<List<RPTChargeDailyEntity>> getChargeDailyList(@RequestBody RPTServicePointWriteOffSearch search) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        List<RPTChargeDailyEntity> list = chargeDailyRptService.getChargeDailyByList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }
}
