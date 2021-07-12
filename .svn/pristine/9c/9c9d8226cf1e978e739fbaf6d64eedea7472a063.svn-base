package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerPraiseDetailsRptService;
import com.kkl.kklplus.provider.rpt.service.KeFuPraiseDetailsRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("customerPraise")
public class CustomerPraiseDetailsRptController {

    @Autowired
    private CustomerPraiseDetailsRptService customerPraiseDetailsRptService;


    @ApiOperation("客户好评明细")
    @PostMapping("getCustomerPraiseDetailsList")
    public MSResponse<MSPage<RPTKeFuPraiseDetailsEntity>> getCustomerPraiseDetailsList(@RequestBody RPTKeFuCompleteTimeSearch search) {
        Page<RPTKeFuPraiseDetailsEntity> pageList = customerPraiseDetailsRptService.getCustomerPraiseDetailsList(search);
        MSPage<RPTKeFuPraiseDetailsEntity> returnPage = new MSPage<>();
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
