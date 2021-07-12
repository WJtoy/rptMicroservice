package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.provider.rpt.service.ServicePointBalanceRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("servicePointBalanceRpt")
public class ServicePointBalanceRptController {

    @Autowired
    private ServicePointBalanceRptService servicePointBalanceRptService;


    @ApiOperation("获取网点余额")
    @PostMapping("getServicePointBalanceByPage")
    public MSResponse<MSPage<RPTServicePointBalanceEntity>> getServicePointBalanceByPage(@RequestBody RPTServicePointWriteOffSearch search) {
        Page<RPTServicePointBalanceEntity>  servicePointList = servicePointBalanceRptService.getServicePointBalanceRptData(search);
        MSPage<RPTServicePointBalanceEntity> returnPage = new MSPage<>();
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
