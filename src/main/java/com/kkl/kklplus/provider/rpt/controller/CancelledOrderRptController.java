package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCancelledOrderSearch;
import com.kkl.kklplus.provider.rpt.service.CancelledOrderRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("cancelledOrder")
public class CancelledOrderRptController {
    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取退单或取消工单明细")
    @PostMapping("getCancelledOrderList")
    public MSResponse<MSPage<RPTCancelledOrderEntity>> getCancelledOrderList(@RequestBody RPTCancelledOrderSearch search) {
        Page<RPTCancelledOrderEntity> pageList = cancelledOrderRptService.getCancelledOrderListByPaging(search);
        MSPage<RPTCancelledOrderEntity> returnPage = new MSPage<>();
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
