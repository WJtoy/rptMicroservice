package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.provider.rpt.service.KeFuPraiseDetailsRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("keFuPraise")
public class KeFuPraiseDetailsRptController {

     @Autowired
     private KeFuPraiseDetailsRptService keFuPraiseDetailsRptService;


    @ApiOperation("客服好评明细")
    @PostMapping("getKeFuPraiseDetailsList")
    public MSResponse<MSPage<RPTKeFuPraiseDetailsEntity>> getCustomerComplainList(@RequestBody RPTKeFuCompleteTimeSearch search) {
        Page<RPTKeFuPraiseDetailsEntity> pageList = keFuPraiseDetailsRptService.getKeFuPraiseDetailsList(search);
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
