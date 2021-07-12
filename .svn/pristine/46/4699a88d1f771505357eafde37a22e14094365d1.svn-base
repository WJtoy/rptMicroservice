package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTServicePointCoverageEntity;
import com.kkl.kklplus.provider.rpt.service.ServicePointCoverageRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("ServicePointCoverage")
public class ServicePointCoverageRptController {

    @Autowired
    private ServicePointCoverageRptService servicePointCoverageRptService;

    /**
     * 条件查询
     */
    @ApiOperation("获取覆盖网点")
    @PostMapping("getServicePointCoverageList")
    public MSResponse<List<RPTServicePointCoverageEntity>> getServicePointCoverageList() {
        List<RPTServicePointCoverageEntity> reminderList = servicePointCoverageRptService.getServicePointCoverAreasRptData();
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }
    /**
     * 条件查询
     */
    @ApiOperation("获取网点没有覆盖区域")
    @PostMapping("getServicePointNoCoverageList")
    public MSResponse<List<RPTServicePointCoverageEntity>> getServicePointNoCoverageList() {
        List<RPTServicePointCoverageEntity> reminderList = servicePointCoverageRptService.getServicePointNoCoverAreasRptData();
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }

}
