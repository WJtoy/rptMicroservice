package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTServicePointPaySummaryEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointPaySummarySearch;
import com.kkl.kklplus.provider.rpt.service.ServicePointChargeRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("servicePointCharge")
public class ServicePointChargeRptController {

    @Autowired
    private ServicePointChargeRptService servicePointChargeRptService;

    @ApiOperation("分页获取网点应付汇总报表")
    @PostMapping("getServicePointPaySummaryRptList")
    public MSResponse<MSPage<RPTServicePointPaySummaryEntity>> getServicePointPaySummaryRptList(@RequestBody RPTServicePointPaySummarySearch search) {
        Page<RPTServicePointPaySummaryEntity> pageList = servicePointChargeRptService.getServicePointPaySummaryRptData(search);
        MSPage<RPTServicePointPaySummaryEntity> returnPage = new MSPage<>();
        if (pageList != null) {
            returnPage.setList(pageList);
            returnPage.setPageNo(pageList.getPageNum());
            returnPage.setPageSize(pageList.getPageSize());
            returnPage.setRowCount((int) pageList.getTotal());
            returnPage.setPageCount(pageList.getPages());
        }
        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }


    @ApiOperation("分页获取网点成本排名报表")
    @PostMapping("getServicePointCostPerRptList")
    public MSResponse<MSPage<RPTServicePointPaySummaryEntity>> getServicePointCostPerRptList(@RequestBody RPTServicePointPaySummarySearch search) {
        Page<RPTServicePointPaySummaryEntity> pageList = servicePointChargeRptService.getServicePointCostPerRptData(search);
        MSPage<RPTServicePointPaySummaryEntity> returnPage = new MSPage<>();
        if (pageList != null) {
            returnPage.setList(pageList);
            returnPage.setPageNo(pageList.getPageNum());
            returnPage.setPageSize(pageList.getPageSize());
            returnPage.setRowCount((int) pageList.getTotal());
            returnPage.setPageCount(pageList.getPages());
        }
        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
