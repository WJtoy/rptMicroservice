package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTServicePointBaseInfoEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointBaseInfoSearch;
import com.kkl.kklplus.provider.rpt.service.ServicePointBaseInfoRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("servicePointBase")
public class ServicePointBaseInfoRptController  {

    @Autowired
    private ServicePointBaseInfoRptService servicePointBaseInfoRptService;


    @ApiOperation("获取基础资料")
    @PostMapping("getServicePointBasePage")
    public MSResponse<MSPage<RPTServicePointBaseInfoEntity>> getServicePointBalanceByPage(@RequestBody RPTServicePointBaseInfoSearch search) {
        MSPage<RPTServicePointBaseInfoEntity> returnPage = servicePointBaseInfoRptService.getServicePointBaseInfoRptDataNew(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
