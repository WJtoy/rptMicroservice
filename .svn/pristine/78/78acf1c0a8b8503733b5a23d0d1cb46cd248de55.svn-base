package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTMasterApplyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.provider.rpt.service.MasterApplyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("masterApply")
public class MasterApplyRptController {

    @Autowired
    private MasterApplyRptService masterApplyRptService;

    @ApiOperation("分页获取配件报表")
    @PostMapping("masterApplyList")
    public MSResponse<MSPage<RPTMasterApplyEntity>> getMasterApplyRptList(@RequestBody RPTComplainStatisticsDailySearch search) {
        Page<RPTMasterApplyEntity> pageList = masterApplyRptService.getMasterApplyList(search);
        MSPage<RPTMasterApplyEntity> returnPage = new MSPage<>();
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
