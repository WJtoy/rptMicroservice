package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BSign;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerSignSearch;
import com.kkl.kklplus.entity.rpt.RPTKeFuAreaEntity;
import com.kkl.kklplus.provider.rpt.service.CustomerContractRptService;
import com.kkl.kklplus.provider.rpt.service.KeFuAreaRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("customerSign")
public class CutsomerContractRptController {

    @Autowired
    private CustomerContractRptService customerContractRptService;

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取客户签约信息")
    @PostMapping("getCustomerSignList")
    public MSResponse<MSPage<B2BSign>> getCompletedOrderList(@RequestBody RPTCustomerSignSearch search) {
        MSPage<B2BSign> pageList = customerContractRptService.getCustomerSignDetailsList(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageList);
    }
}
