package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCrushCoverageEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuAreaEntity;
import com.kkl.kklplus.provider.rpt.service.CrushCoverageRptService;
import com.kkl.kklplus.provider.rpt.service.KeFuAreaRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("keFuArea")
public class KeFuAreaRptController {

    @Autowired
    private KeFuAreaRptService keFuAreaRptService;

    /**
     * 条件查询
     */
    @ApiOperation("获取客服")
    @PostMapping("getKeFuAreaList")
    public MSResponse<List<RPTKeFuAreaEntity>> getKeFuAreaList() {
        List<RPTKeFuAreaEntity> reminderList = keFuAreaRptService.getKeFuAreasRptData();
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }
}
