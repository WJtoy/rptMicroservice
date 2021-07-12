package com.kkl.kklplus.provider.rpt.controller;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;

import com.kkl.kklplus.entity.rpt.RPTSpecialChargeAreaEntity;
import com.kkl.kklplus.entity.rpt.search.RPTSpecialChargeSearchCondition;


import com.kkl.kklplus.provider.rpt.service.SpecialChargeAreaRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.List;


@RestController
@RequestMapping("specialChargeArea")
public class SpecialChargeAreaRptController {
    @Autowired
    private SpecialChargeAreaRptService specialChargeAreaRptService;

    /**
     * 查询特殊费用分布
     *
     * @return
     */
    @ApiOperation("条件(areaId,yearmonth,)获取特殊费用分布信息")
    @PostMapping("getList")
    public MSResponse<List<RPTSpecialChargeAreaEntity>> getSpecialChargeAreaList(@RequestBody RPTSpecialChargeSearchCondition rptSearchCondtion
                                                                                            ) {
        List<RPTSpecialChargeAreaEntity> specialChargeNewList = specialChargeAreaRptService.getSpecialChargeNewList(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, specialChargeNewList);
    }



}
