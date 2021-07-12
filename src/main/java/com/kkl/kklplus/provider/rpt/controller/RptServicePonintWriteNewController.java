package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.ServicePointChargeRptEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.provider.rpt.service.RptServicePonintWriteNewService;
import com.kkl.kklplus.provider.rpt.service.RptServicePonintWriteService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("rptServicePointWriteNew")
public class RptServicePonintWriteNewController {

    @Autowired
    private RptServicePonintWriteNewService rptServicePonintWriteNewService;

    /**
     * 查询网点对账明细(新)
     *
     * @return
     */
    @ApiOperation("获取网点对账")
    @PostMapping("getNrPointWriteOffNew")
    public MSResponse<MSPage<ServicePointChargeRptEntity>> getNrPointWriteOff(@RequestBody RPTServicePointWriteOffSearch search) {
        Page<ServicePointChargeRptEntity> pointWriteOffList = rptServicePonintWriteNewService.getServiceWriteRptList(search);
        MSPage<ServicePointChargeRptEntity> returnPage = new MSPage<>();
        if (pointWriteOffList != null) {
            returnPage.setList(pointWriteOffList);
            returnPage.setPageNo(pointWriteOffList.getPageNum());
            returnPage.setPageSize(pointWriteOffList.getPageSize());
            returnPage.setRowCount((int) pointWriteOffList.getTotal());
            returnPage.setPageCount(pointWriteOffList.getPages());
        }

        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }
}
