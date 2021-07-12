package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCrushCoverageEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointCoverageEntity;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.provider.rpt.service.CrushCoverageRptService;
import com.kkl.kklplus.provider.rpt.service.ServicePointCoverageRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("crushCoverage")
public class CrushCoverageRptController {

    @Autowired
    private CrushCoverageRptService crushCoverageRptService;

    /**
     * 条件查询
     */
    @ApiOperation("获取突击区域")
    @PostMapping("getCrushCoverageList")
    public MSResponse<List<RPTCrushCoverageEntity>> getServicePointCoverageList(@RequestBody RPTGradedOrderSearch rptGradedOrderSearch) {
        List<RPTCrushCoverageEntity> reminderList = crushCoverageRptService.getCrushCoverAreasRptData(rptGradedOrderSearch);
        return new MSResponse<>(MSErrorCode.SUCCESS, reminderList);
    }
}
