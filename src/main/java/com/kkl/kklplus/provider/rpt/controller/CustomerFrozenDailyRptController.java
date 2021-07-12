package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerFrozenDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerNewOrderDailyRptEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.provider.rpt.service.CustomerFrozenDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("customerFrozenDaily")
public class CustomerFrozenDailyRptController {


    @Autowired
    private CustomerFrozenDailyRptService customerFrozenDailyRptService;

    /**
     * 查询每日冻结明细列表
     *
     * @return
     */
    @ApiOperation("条件(customerId,salesId,beginDate*)获取客户每日冻结明细数据分页")
    @PostMapping("getCustomerFrozenDailyRptList")
    public MSResponse<MSPage<RPTCustomerFrozenDailyEntity>> getCustomerNewOrderDailyList(@RequestBody RPTSearchCondtion rptSearchCondtion) {
        Page<RPTCustomerFrozenDailyEntity> customerFrozenDailyDailyPage = customerFrozenDailyRptService.getCustomerFrozenDailyList(rptSearchCondtion);
        MSPage<RPTCustomerFrozenDailyEntity> pageRecord = new MSPage<>();

       if(customerFrozenDailyDailyPage != null) {
           pageRecord.setPageNo(customerFrozenDailyDailyPage.getPageNum());
           pageRecord.setPageSize(customerFrozenDailyDailyPage.getPageSize());
           pageRecord.setRowCount((int) customerFrozenDailyDailyPage.getTotal());
           pageRecord.setPageCount(customerFrozenDailyDailyPage.getPages());
           pageRecord.setList(customerFrozenDailyDailyPage);
       }

         return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }
}
