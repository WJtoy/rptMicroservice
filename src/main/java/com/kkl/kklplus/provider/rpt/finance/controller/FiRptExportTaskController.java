package com.kkl.kklplus.provider.rpt.finance.controller;

import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTExportTaskEntity;
import com.kkl.kklplus.entity.rpt.RPTExportTaskSearch;
import com.kkl.kklplus.entity.rpt.common.RPTErrorCode;
import com.kkl.kklplus.entity.rpt.common.RPTReportEnum;
import com.kkl.kklplus.entity.rpt.common.RPTUserTypeEnum;
import com.kkl.kklplus.entity.rpt.exception.RPTBaseException;
import com.kkl.kklplus.provider.rpt.service.RptExportTaskService;
import io.swagger.annotations.ApiOperation;
import org.javatuples.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("finance/rptExportTask")
public class FiRptExportTaskController {

    @Autowired
    private RptExportTaskService rptExportTaskService;

    @ApiOperation("检查导出任务(参数：reportId、taskCreateBy、searchConditionJson)")
    @PostMapping("checkRptExportTask")
    public MSResponse<String> checkRptExportTask(@RequestBody RPTExportTaskEntity taskEntity) {
        MSErrorCode errorCode = MSErrorCode.SUCCESS;
        RPTReportEnum reportEnum = RPTReportEnum.valueOf(taskEntity.getReportId());
        if (reportEnum != null && RPTUserTypeEnum.FINANCE.equals(reportEnum.userType) && taskEntity.getTaskCreateBy() != null) {
            try {
                rptExportTaskService.checkRptExportTask(taskEntity);
            } catch (RPTBaseException e) {
                errorCode = e.getErrorCode();
            } catch (Exception e1) {
                errorCode = new MSErrorCode(RPTErrorCode.RPT_OPERATE_FAILURE.getCode(), "导出失败：" + e1.getLocalizedMessage());
            }
            return new MSResponse<>(errorCode, "");
        } else {
            return new MSResponse<>(RPTErrorCode.RPT_NO_AUTHORIZATION, "");
        }
    }

    @ApiOperation("创建报表导出任务(参数：reportId、reportType、taskCreateBy、taskCreateByName、reportTitle、searchConditionJson)")
    @PostMapping("createRptExportTask")
    public MSResponse<String> createRptExportTask(@RequestBody RPTExportTaskEntity taskEntity) {
        RPTReportEnum reportEnum = RPTReportEnum.valueOf(taskEntity.getReportId());
        if (reportEnum != null && RPTUserTypeEnum.FINANCE.equals(reportEnum.userType) && taskEntity.getTaskCreateBy() != null) {
            MSErrorCode errorCode = MSErrorCode.SUCCESS;
            try {
                rptExportTaskService.createRptExportTask(taskEntity);
            } catch (RPTBaseException e) {
                errorCode = e.getErrorCode();
            } catch (Exception e1) {
                errorCode = new MSErrorCode(RPTErrorCode.RPT_OPERATE_FAILURE.getCode(), "创建报表导出任务失败：" + e1.getLocalizedMessage());
            }
            return new MSResponse<>(errorCode, "");
        } else {
            return new MSResponse<>(RPTErrorCode.RPT_NO_AUTHORIZATION, "");
        }
    }

    /**
     * 分页&条件查询
     */
    @ApiOperation("分页获取报表导出任务列表")
    @PostMapping("getRptExportTaskList")
    public MSResponse<MSPage<RPTExportTaskEntity>> getRptExportTaskList(@RequestBody RPTExportTaskSearch search) {
        Page<RPTExportTaskEntity> pageList = rptExportTaskService.getRptExportTaskList(search);
        MSPage<RPTExportTaskEntity> pageRecord = new MSPage<>();
        if (pageList != null) {
            pageRecord.setList(pageList);
            pageRecord.setPageNo(pageList.getPageNum());
            pageRecord.setPageSize(pageList.getPageSize());
            pageRecord.setRowCount((int) pageList.getTotal());
            pageRecord.setPageCount(pageList.getPages());
        }
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);
    }


    /**
     * 获取报表下载地址
     */
    @ApiOperation("获取报表Excel文件下载地址")
    @PostMapping("getRptExcelDownloadUrl")
    public MSResponse<String> getRptExcelDownloadUrl(@RequestBody RPTExportTaskEntity taskEntity) {
        MSResponse<String> response = null;
        try {
            RPTReportEnum reportEnum = RPTReportEnum.valueOf(taskEntity.getReportId());
            if (reportEnum != null && RPTUserTypeEnum.FINANCE.equals(reportEnum.userType)) {
                Pair<Boolean, String> result = rptExportTaskService.downloadReportExcel(taskEntity);
                if (result.getValue0()) {
                    response = new MSResponse<>(MSErrorCode.SUCCESS, result.getValue1());
                }
            } else {
                return new MSResponse<>(RPTErrorCode.RPT_NO_AUTHORIZATION, "");
            }
        } catch (Exception e) {
            response = new MSResponse<>(MSErrorCode.FAILURE, "");
        }
        return response;
    }

}
