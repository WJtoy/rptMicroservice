package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerNewOrderDailyRptEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.provider.rpt.service.CustomerOrderDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("customerNewOrderDailyRpt")
public class CustomerNewOrderDailyRptController {
    @Autowired
    private CustomerOrderDailyRptService customerOrderDailyRptService;

    /**
     * 查询客户每日下单列表
     *
     * @return
     */
    @ApiOperation("条件(customerId,salesId,beginDate*)获取客户每日下单明细数据分页")
    @PostMapping("getList")
    public MSResponse<MSPage<RPTCustomerNewOrderDailyRptEntity>> getCustomerNewOrderDailyList(@RequestBody RPTSearchCondtion rptSearchCondtion) {
        Page<RPTCustomerNewOrderDailyRptEntity> customerNewOrderDailyPage = customerOrderDailyRptService.getCustomerNewOrderDailyList(rptSearchCondtion);
        MSPage<RPTCustomerNewOrderDailyRptEntity> pageRecord = new MSPage<>();
        pageRecord.setPageNo(customerNewOrderDailyPage.getPageNum());
        pageRecord.setPageSize(customerNewOrderDailyPage.getPageSize());
        pageRecord.setRowCount((int) customerNewOrderDailyPage.getTotal());
        pageRecord.setPageCount(customerNewOrderDailyPage.getPages());
        pageRecord.setList(customerNewOrderDailyPage);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }
}
