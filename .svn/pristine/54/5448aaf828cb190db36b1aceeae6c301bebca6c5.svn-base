package com.kkl.kklplus.provider.rpt.controller;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.provider.rpt.service.ComplainRatioDailyRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("complainOrder")
public class ComplainRatioDailyRptController {

    @Autowired
    private ComplainRatioDailyRptService complainRatioDailyRptService;

    /**
     * 查询省每日完工
     *
     * @return
     */
    @ApiOperation("获取省每日完工数据")
    @PostMapping("provinceComplainCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getProvinceComplainCompletedList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> provinceComplainCompletedOrderData = complainRatioDailyRptService.getProvinceCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, provinceComplainCompletedOrderData);
    }

    /**
     * 查询市每日完工
     *
     * @return
     */
    @ApiOperation("获取市每日完工数据")
    @PostMapping("cityComplainCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getCityComplainCompletedList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> cityComplainCompletedOrderData = complainRatioDailyRptService.getCityCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, cityComplainCompletedOrderData);
    }

    /**
     * 查询区每日完工
     *
     * @return
     */
    @ApiOperation("获取区每日完工数据")
    @PostMapping("areaComplainCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getAreaComplainCompletedList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> areaComplainCompletedOrderData = complainRatioDailyRptService.getAreaCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, areaComplainCompletedOrderData);
    }


}
