package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import com.kkl.kklplus.entity.rpt.RPTReminderResponseTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerReminderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTReminderResponseTimeSearch;
import com.kkl.kklplus.provider.rpt.service.CustomerReminderRptService;
import com.kkl.kklplus.provider.rpt.service.KaReminderResponseTimeRptService;
import com.kkl.kklplus.provider.rpt.service.ReminderResponseTimeRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("kaReminder")
public class KaReminderRptController {

    @Autowired
    private KaReminderResponseTimeRptService kaReminderResponseTimeRptService;

    @ApiOperation("分页获取催单回复时效")
    @PostMapping("getKaReminderResponseTimeList")
    public MSResponse<MSPage<RPTReminderResponseTimeEntity>> getReminderResponseTimeList(@RequestBody RPTReminderResponseTimeSearch search) {
        Page<RPTReminderResponseTimeEntity> pageList = kaReminderResponseTimeRptService.getReminderResponseTimeRptData(search);
        MSPage<RPTReminderResponseTimeEntity> returnPage = new MSPage<>();
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
