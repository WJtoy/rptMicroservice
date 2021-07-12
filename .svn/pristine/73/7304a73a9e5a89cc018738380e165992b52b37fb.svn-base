package com.kkl.kklplus.provider.rpt.controller;


import com.github.pagehelper.Page;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTTravelChargeRankEntity;
import com.kkl.kklplus.entity.rpt.search.RPTTravelChargeRankSearchCondition;
import com.kkl.kklplus.provider.rpt.service.TravelChargeRankRptService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
@RequestMapping("travelChargeRank")
public class TravelChargeRankRptController {
    @Autowired
    private TravelChargeRankRptService travelChargeRankRptService;

    /**
     * 查询远程费用排名
     *
     * @return
     */
    @ApiOperation("获取远程费用排名数据分页")
    @PostMapping("getList")
    public MSResponse<MSPage<RPTTravelChargeRankEntity>> getTravelChargeRank(@RequestBody RPTTravelChargeRankSearchCondition rptSearchCondtion) {
        Page<RPTTravelChargeRankEntity> travelChargeRankList = travelChargeRankRptService.getTravelChargeRankList(rptSearchCondtion);
        MSPage<RPTTravelChargeRankEntity> pageRecord = new MSPage<>();
        pageRecord.setPageNo(travelChargeRankList.getPageNum());
        pageRecord.setPageSize(travelChargeRankList.getPageSize());
        pageRecord.setPageCount(travelChargeRankList.getPages());
        pageRecord.setRowCount((int) travelChargeRankList.getTotal());
        pageRecord.setList(travelChargeRankList);
        return new MSResponse<>(MSErrorCode.SUCCESS, pageRecord);

    }


}
