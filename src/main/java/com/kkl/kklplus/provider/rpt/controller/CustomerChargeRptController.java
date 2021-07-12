package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCancelledOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerWriteOffSearch;
import com.kkl.kklplus.provider.rpt.service.*;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("customerCharge")
public class CustomerChargeRptController {
    @Autowired
    private CompletedOrderRptService completedOrderRptService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;

    @Autowired
    private CustomerChargeSummaryRptService customerChargeSummaryRptService;

    @Autowired
    private CustomerChargeSummaryRptNewService customerChargeSummaryRptNewService;


    /**
     * 获取客户的工单数量与消费金额信息
     */
    @ApiOperation("获取客户的工单数量与消费金额信息")
    @PostMapping("getCustomerChargeSummaryMonthly")
    public MSResponse<RPTCustomerChargeSummaryMonthlyEntity> getCustomerChargeSummaryMonthly(@RequestBody RPTCustomerChargeSearch search) {
        RPTCustomerChargeSummaryMonthlyEntity entity = customerChargeSummaryRptNewService.getCustomerChargeSummaryNew(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, entity);
    }

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获完工单明细")
    @PostMapping("getCompletedOrderList")
    public MSResponse<MSPage<RPTCompletedOrderEntity>> getCompletedOrderList(@RequestBody RPTCompletedOrderSearch search) {
        Page<RPTCompletedOrderEntity> pageList = completedOrderRptService.getCompletedOrderListByPaging(search);
        MSPage<RPTCompletedOrderEntity> returnPage = new MSPage<>();
        if (pageList != null) {
            returnPage.setList(pageList);
            returnPage.setPageNo(pageList.getPageNum());
            returnPage.setPageSize(pageList.getPageSize());
            returnPage.setRowCount((int) pageList.getTotal());
            returnPage.setPageCount(pageList.getPages());
        }
        return new MSResponse<>(MSErrorCode.SUCCESS, returnPage);
    }


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
    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取客户退补明细")
    @PostMapping("getCustomerWriteOffList")
    public MSResponse<MSPage<RPTCustomerWriteOffEntity>> getCustomerWriteOffList(@RequestBody RPTCustomerWriteOffSearch search) {
        Page<RPTCustomerWriteOffEntity> pageList = customerWriteOffRptService.getCustomerWriteOffListByPaging(search);
        MSPage<RPTCustomerWriteOffEntity> returnPage = new MSPage<>();
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
