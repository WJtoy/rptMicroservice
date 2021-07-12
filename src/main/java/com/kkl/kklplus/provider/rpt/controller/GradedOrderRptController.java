package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;

import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTDevelopAverageOrderFeeEntity;
import com.kkl.kklplus.entity.rpt.RPTGradedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTKefuCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.provider.rpt.service.GradedOrderRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;


@RestController
@RequestMapping("gradedOrder")
public class GradedOrderRptController {
    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    /**
     * 查询工单费用报表数据
     *
     * @return
     */
    @ApiOperation("获取工单费用报表数据分页")
    @PostMapping("orderServicePointFee")
    public MSResponse<MSPage<RPTGradedOrderEntity>> getOrderServicePointFee(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        Page<RPTGradedOrderEntity> gradedOrderList = gradedOrderRptService.getOrderServicePointFeeOfGradedOrder(rptSearchCondtion);
        MSPage<RPTGradedOrderEntity> pageRecord = new MSPage<>();
        pageRecord.setPageNo(gradedOrderList.getPageNum());
        pageRecord.setPageSize(gradedOrderList.getPageSize());
        pageRecord.setPageCount(gradedOrderList.getPages());
        pageRecord.setRowCount((int) gradedOrderList.getTotal());
        pageRecord.setList(gradedOrderList);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }

    /**
     * 查询客服每日完工
     *
     * @return
     */
    @ApiOperation("获取客服每日完工数据")
    @PostMapping("kefuDailyCompleted")
    public MSResponse<List<RPTKefuCompletedDailyEntity>> getkefuCompletedDailyList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTKefuCompletedDailyEntity> kefuCompletedDaily = gradedOrderRptService.getKefuCompletedDaily(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, kefuCompletedDaily);
    }

    /**
     * 查询省每日完工
     *
     * @return
     */
    @ApiOperation("获取省每日完工数据")
    @PostMapping("provinceCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getprovinceCompletedOrderList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> provinceCompletedOrderData = gradedOrderRptService.getProvinceCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, provinceCompletedOrderData);
    }

    /**
     * 查询市每日完工
     *
     * @return
     */
    @ApiOperation("获取市每日完工数据")
    @PostMapping("cityCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getCityCompletedOrderList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> cityCompletedOrderData = gradedOrderRptService.getCityCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, cityCompletedOrderData);
    }

    /**
     * 查询区每日完工
     *
     * @return
     */
    @ApiOperation("获取区每日完工数据")
    @PostMapping("areaCompleteOrder")
    public MSResponse<List<RPTAreaCompletedDailyEntity>> getAreaCompletedOrderList(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTAreaCompletedDailyEntity> areaCompletedOrderData = gradedOrderRptService.getAreaCompletedOrderData(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, areaCompletedOrderData);
    }

    /**
     * 查询开发均单费用
     *
     * @return
     */
    @ApiOperation("获取开发均单费用数据")
    @PostMapping("developAverageFee")
    public MSResponse<List<RPTDevelopAverageOrderFeeEntity>> getDevelopAverageOrderFee(@RequestBody RPTGradedOrderSearch rptSearchCondtion) {
        List<RPTDevelopAverageOrderFeeEntity> list = gradedOrderRptService.getDevelopAverageOrderFee(rptSearchCondtion);
        return new MSResponse<>(MSErrorCode.SUCCESS, list);
    }



}
