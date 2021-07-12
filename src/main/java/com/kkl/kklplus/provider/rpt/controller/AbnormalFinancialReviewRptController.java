package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTAbnormalFinancialAuditEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.AbnormalFinancialReviewRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("abnormalFinancial")
public class AbnormalFinancialReviewRptController {

    @Autowired
    private AbnormalFinancialReviewRptService abnormalFinancialReviewRptService;

    /**
     * 获取每日财务审单
     *
     * @return
     */
    @ApiOperation("获取每日财务审单")
    @PostMapping("getAbnormalFinancialList")
    public MSResponse<List<RPTAbnormalFinancialAuditEntity>> getAbnormalFinancialList(@RequestBody RPTCustomerOrderPlanDailySearch rptSearchCondtion) {
        List<RPTAbnormalFinancialAuditEntity> list = abnormalFinancialReviewRptService.getAbnormalFinancialAuditList(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }

}
