package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.service.CompletedOrderDetailsService;
import com.kkl.kklplus.provider.rpt.service.CompletedOrderNewDetailsService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("completedOrderNew")
public class CompletedOrderNewRptController {
    @Autowired
    private CompletedOrderNewDetailsService completedOrderNewDetailsService;

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取完工单明细(财务专用)")
    @PostMapping("getCompletedOrderNewDetailsList")
    public MSResponse<MSPage<RPTCompletedOrderDetailsEntity>> getCompletedOrderList(@RequestBody RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> pageList = completedOrderNewDetailsService.getCompletedOrderNewDetailsRptList(search);
        MSPage<RPTCompletedOrderDetailsEntity> returnPage = new MSPage<>();
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
