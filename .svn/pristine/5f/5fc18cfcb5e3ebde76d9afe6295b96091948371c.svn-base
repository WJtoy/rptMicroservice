package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerFinanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerFinanceSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerFinanceRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("customerFinance")
public class CustomerFinanceRptController {

    @Autowired
    private CustomerFinanceRptService customerFinanceRptService;

    @ApiOperation("分页获取报表")
    @PostMapping("getCustomerFinanceRptList")
    public MSResponse<MSPage<RPTCustomerFinanceEntity>> getCustomerFinanceRptList(@RequestBody RPTCustomerFinanceSearch search) {
        MSPage<RPTCustomerFinanceEntity> pageList = customerFinanceRptService.getCustomerFinanceData(search,null);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageList);
    }
}
