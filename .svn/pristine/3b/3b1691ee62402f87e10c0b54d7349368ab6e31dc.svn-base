package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTKeFuAverageOrderFeeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.provider.rpt.service.KeFuAverageOrderFeeRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("keFuAverageOrderFee")
public class KeFuAverageOrderFeeRptController {

    @Autowired
    private KeFuAverageOrderFeeRptService keFuAverageOrderFeeRptService;


    @ApiOperation("获取非KA其他费用")
    @PostMapping("keFuAverageOrderFeeList")
    public MSResponse<List<RPTKeFuAverageOrderFeeEntity>> getkeFuAverageOrderFee(@RequestBody RPTComplainStatisticsDailySearch search) {
        List<RPTKeFuAverageOrderFeeEntity> returnList = keFuAverageOrderFeeRptService.getKeFuAverageOrderFee(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);
    }



    @ApiOperation("获取KA其他费用")
    @PostMapping("vipKeFuAverageOrderFeeList")
    public MSResponse<List<RPTKeFuAverageOrderFeeEntity>> getVipKeFuAverageOrderFee(@RequestBody RPTComplainStatisticsDailySearch search) {
        List<RPTKeFuAverageOrderFeeEntity> returnList = keFuAverageOrderFeeRptService.getKAKeFuAverageOrderFee(search);
        return new MSResponse<>(MSErrorCode.SUCCESS, returnList);
    }
}
