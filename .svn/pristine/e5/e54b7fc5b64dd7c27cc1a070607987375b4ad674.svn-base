package com.kkl.kklplus.provider.rpt.controller;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTCrushAreaEntity;
import com.kkl.kklplus.entity.rpt.RPTSpecialChargeAreaEntity;
import com.kkl.kklplus.entity.rpt.search.RPTSpecialChargeSearchCondition;
import com.kkl.kklplus.provider.rpt.service.CrushAreaRptService;
import com.kkl.kklplus.provider.rpt.service.SpecialChargeAreaRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;


@RestController
@RequestMapping("crushArea")
public class CrushAreaRptController {
    @Autowired
    private CrushAreaRptService crushAreaRptService;

    /**
     * 查询特殊费用分布
     *
     * @return
     */
    @ApiOperation("突击单量")
    @PostMapping("getCrushList")
    public MSResponse<List<RPTCrushAreaEntity>> getCrushList(@RequestBody RPTSpecialChargeSearchCondition rptSearchCondtion
                                                                                            ) {
        List<RPTCrushAreaEntity> specialChargeNewList = crushAreaRptService.getCrushAreaData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, specialChargeNewList);
    }



}
