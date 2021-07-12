package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTUncompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTUncompletedQtyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTUncompletedOrderSearch;
import com.kkl.kklplus.provider.rpt.service.UncompletedOrderRptService;
import com.kkl.kklplus.provider.rpt.service.UncompletedQtyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("uncompletedOrderNew")
public class UncompletedQtyRptController {

    @Autowired
    private UncompletedQtyRptService uncompletedQtyRptService;


    @ApiOperation("分页获未完工单数量")
    @PostMapping("getUncompletedOrderNewList")
    public MSResponse<MSPage<RPTUncompletedQtyEntity>> getUncompletedOrderList(@RequestBody RPTUncompletedOrderSearch search) {
        Page<RPTUncompletedQtyEntity> pageList = uncompletedQtyRptService.getUncompletedQtyRptData(search);
        MSPage<RPTUncompletedQtyEntity> returnPage = new MSPage<>();
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
