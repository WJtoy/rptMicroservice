package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTFinancialReviewDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.service.FinancialReviewDetailsRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("financialReviewDetails")
public class FinancialReviewDetailsRptController {

    @Autowired
    private FinancialReviewDetailsRptService financialReviewDetailsRptService;


    @ApiOperation("财务审单明细")
    @PostMapping("getFinancialReviewDetailsList")
    public MSResponse<MSPage<RPTFinancialReviewDetailsEntity>> getFinancialReviewDetailsList(@RequestBody RPTCustomerOrderPlanDailySearch rptSearchCondtion) {
        Page<RPTFinancialReviewDetailsEntity> financialReviewDetailsPage = financialReviewDetailsRptService.getFinancialReviewList(rptSearchCondtion);
        MSPage<RPTFinancialReviewDetailsEntity> pageRecord = new MSPage<>();

        if(financialReviewDetailsPage != null) {
            pageRecord.setPageNo(financialReviewDetailsPage.getPageNum());
            pageRecord.setPageSize(financialReviewDetailsPage.getPageSize());
            pageRecord.setRowCount((int) financialReviewDetailsPage.getTotal());
            pageRecord.setPageCount(financialReviewDetailsPage.getPages());
            pageRecord.setList(financialReviewDetailsPage);
        }

        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }
}
