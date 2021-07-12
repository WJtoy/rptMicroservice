package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerComplainSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerComplainRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("customerComplain")
public class CustomerComplainRptController {

    @Autowired
    private CustomerComplainRptService customerComplainRptService;

    @ApiOperation("分页获客户投诉明细")
    @PostMapping("getCustomerComplainList")
    public MSResponse<MSPage<RPTCustomerComplainEntity>> getCustomerComplainList(@RequestBody RPTCustomerComplainSearch search) {
        Page<RPTCustomerComplainEntity> pageList = customerComplainRptService.getCustomerComplainNewList(search);
        MSPage<RPTCustomerComplainEntity> returnPage = new MSPage<>();
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
