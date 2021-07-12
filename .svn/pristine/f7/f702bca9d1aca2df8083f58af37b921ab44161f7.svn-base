package com.kkl.kklplus.provider.rpt.servicepoint.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.ServicePointChargeRptEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
import com.kkl.kklplus.provider.rpt.service.RptServicePonintWriteService;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("ServciePointReconciliation")
public class ServciePointReconciliationPptController extends RptBaseService {

    @Autowired
    private RptServicePonintWriteService rptServicePonintWriteService;

    /**
     * 查询网点对账明细
     *
     * @return
     */
    @ApiOperation("获取网点对账")
    @PostMapping("getServciePointReconciliation")
    public MSResponse<MSPage<ServicePointChargeRptEntity>> getNrPointWriteOff(@RequestBody RPTServicePointWriteOffSearch search) {
        Page<ServicePointChargeRptEntity> pointWriteOffList = rptServicePonintWriteService.getServiceWriteRptList(search);
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
