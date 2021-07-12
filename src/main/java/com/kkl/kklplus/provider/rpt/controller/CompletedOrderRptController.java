package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.service.CompletedOrderDetailsService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("completedOrder")
public class CompletedOrderRptController {
    @Autowired
    private CompletedOrderDetailsService completedOrderDetailsService;

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取完工单明细")
    @PostMapping("getCompletedOrderDetailsList")
    public MSResponse<MSPage<RPTCompletedOrderDetailsEntity>> getCompletedOrderList(@RequestBody RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> pageList = completedOrderDetailsService.getCompletedOrderDetailsRptList(search);
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
