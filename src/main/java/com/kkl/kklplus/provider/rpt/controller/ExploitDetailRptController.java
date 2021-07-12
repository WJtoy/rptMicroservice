package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTExploitDetailEntity;
import com.kkl.kklplus.entity.rpt.search.RPTExploitDetailSearch;
import com.kkl.kklplus.provider.rpt.service.ExploitDetailRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("exploitDetail")
public class ExploitDetailRptController {
    @Autowired
    private ExploitDetailRptService exploitDetailRptService;


    @ApiOperation("分页获取开发明细报表")
    @PostMapping("getExploitDetailRptList")
    public MSResponse<MSPage<RPTExploitDetailEntity>> getExploitDetailRptList(@RequestBody RPTExploitDetailSearch search) {
        Page<RPTExploitDetailEntity> pageList = exploitDetailRptService.getExploitDetailRptData(search);
        MSPage<RPTExploitDetailEntity> returnPage = new MSPage<>();
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
