package com.kkl.kklplus.provider.rpt.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTRebuildMiddleTableTaskEntity;
import com.kkl.kklplus.entity.rpt.common.RPTErrorCode;
import com.kkl.kklplus.entity.rpt.exception.RPTBaseException;
import com.kkl.kklplus.entity.rpt.search.RPTRebuildMiddleTableTaskSearch;
import com.kkl.kklplus.provider.rpt.service.RebuildMiddleTableService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("rebuildMiddleTableTask")
public class RebuildMiddleTableTaskController {

    @Autowired
    private RebuildMiddleTableService rebuildMiddleTableService;

    @ApiOperation("创建报表中间表重建任务")
    @PostMapping("createRebuildMiddleTableTask")
    public MSResponse<String> createRebuildMiddleTableTask(@RequestBody RPTRebuildMiddleTableTaskEntity taskEntity) {
        MSErrorCode errorCode = MSErrorCode.SUCCESS;
        try {
            rebuildMiddleTableService.createRebuildMiddleTableTask(taskEntity);
        } catch (RPTBaseException e) {
            errorCode = e.getErrorCode();
        } catch (Exception e1) {
            errorCode = new MSErrorCode(RPTErrorCode.RPT_OPERATE_FAILURE.getCode(), "重建中间表数据失败：" + e1.getLocalizedMessage());
        }
        return new MSResponse<>(errorCode, "");
    }

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取报表中间表重建任务列表")
    @PostMapping("getRebuildMiddleTableTaskList")
    public MSResponse<MSPage<RPTRebuildMiddleTableTaskEntity>> getRebuildMiddleTableTaskList(@RequestBody RPTRebuildMiddleTableTaskSearch search) {
        Page<RPTRebuildMiddleTableTaskEntity> pageList = rebuildMiddleTableService.getRebuildMiddleTableTaskList(search);
        MSPage<RPTRebuildMiddleTableTaskEntity> pageRecord = new MSPage<>();
        if (pageList != null) {
            pageRecord.setList(pageList);
            pageRecord.setPageNo(pageList.getPageNum());
            pageRecord.setPageSize(pageList.getPageSize());
            pageRecord.setRowCount((int) pageList.getTotal());
            pageRecord.setPageCount(pageList.getPages());
        }
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);
    }


}
