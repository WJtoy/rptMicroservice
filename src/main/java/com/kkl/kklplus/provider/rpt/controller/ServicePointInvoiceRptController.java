package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointInvoiceSearch;
import com.kkl.kklplus.provider.rpt.service.ServicePointInvoiceRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("servicePointInvoice")
public class ServicePointInvoiceRptController {

    @Autowired
    private ServicePointInvoiceRptService getServicePointInvoiceList;

    @ApiOperation("获取网点付款清单")
    @PostMapping("getServicePointInvoiceList")
    public MSResponse<MSPage<RPTServicePointInvoiceEntity>> getServicePointInvoiceList(@RequestBody RPTServicePointInvoiceSearch search) {
        Page<RPTServicePointInvoiceEntity> servicePointList = getServicePointInvoiceList.getServicePointInvoiceRptPage(search);
        MSPage<RPTServicePointInvoiceEntity> returnPage = new MSPage<>();
        if (servicePointList != null) {
            returnPage.setList(servicePointList);
            returnPage.setPageNo(servicePointList.getPageNum());
            returnPage.setPageSize(servicePointList.getPageSize());
            returnPage.setRowCount((int) servicePointList.getTotal());
            returnPage.setPageCount(servicePointList.getPages());
        }

        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
