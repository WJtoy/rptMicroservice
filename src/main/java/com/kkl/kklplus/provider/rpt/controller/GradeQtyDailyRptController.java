package com.kkl.kklplus.provider.rpt.controller;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTGradeQtyDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.GradeQtyDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("gradeQtyDaily")
public class GradeQtyDailyRptController {

    @Autowired
    private GradeQtyDailyRptService gradeQtyDailyRptService;

    @ApiOperation("获取客评统计")
    @PostMapping("gradeQtyDailyByList")
    public MSResponse<List<RPTGradeQtyDailyEntity>> gradeQtyDailyByList(@RequestBody RPTCustomerOrderPlanDailySearch search) {
        List<RPTGradeQtyDailyEntity> kefuCompletedDaily = gradeQtyDailyRptService.getGradeQtyRptList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, kefuCompletedDaily);
    }
}
