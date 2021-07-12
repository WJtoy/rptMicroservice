package com.kkl.kklplus.provider.rpt.customer.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerNewOrderDailyRptEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.provider.rpt.customer.service.CtCustomerOrderMonthRptService;
import com.kkl.kklplus.provider.rpt.service.CustomerOrderMonthRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("customer/customerNewOrderMonthRpt")
public class CtCustomerNewOrderMonthRptController {
    @Autowired
    private CtCustomerOrderMonthRptService ctCustomerOrderMonthRptService;

    /**
     * 查询客户每月下单列表
     *
     * @return
     */
    @ApiOperation("条件(customerId,salesId,beginDate*)获取客户每月下单明细数据分页")
    @PostMapping("getMonthOrderList")
    public MSResponse<MSPage<RPTCustomerNewOrderDailyRptEntity>> getCustomerNewOrderDailyList(@RequestBody RPTSearchCondtion rptSearchCondtion) {
        Page<RPTCustomerNewOrderDailyRptEntity> customerNewOrderDailyPage = ctCustomerOrderMonthRptService.getCustomerNewOrderMonthList(rptSearchCondtion);
        MSPage<RPTCustomerNewOrderDailyRptEntity> pageRecord = new MSPage<>();
        pageRecord.setPageNo(customerNewOrderDailyPage.getPageNum());
        pageRecord.setPageSize(customerNewOrderDailyPage.getPageSize());
        pageRecord.setRowCount((int) customerNewOrderDailyPage.getTotal());
        pageRecord.setPageCount(customerNewOrderDailyPage.getPages());
        pageRecord.setList(customerNewOrderDailyPage);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }
}
