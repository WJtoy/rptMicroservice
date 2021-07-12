package com.kkl.kklplus.provider.rpt.chart.service;

import com.kkl.kklplus.entity.rpt.RPTCustomerComplainChartEntity;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.mapper.CustomerComplainChartMapper;
import com.kkl.kklplus.provider.rpt.chart.mapper.KeFuCompleteTimeChartMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerComplainChartService {

    @Resource
    private CustomerComplainChartMapper customerComplainChartMapper;

    public RPTCustomerComplainChartEntity getCustomerComplain(RPTDataDrawingListSearch search) {
        RPTCustomerComplainChartEntity entity = new RPTCustomerComplainChartEntity();
        int systemId = RptCommonUtils.getSystemId();
        Long endDate = DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime();
        Long startDate = DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime();
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        Integer customerComplainQty = customerComplainChartMapper.getCustomerComplainQty(systemId, startDate, endDate, quarter);
        Integer customerValidComplain = customerComplainChartMapper.getCustomerValidComplain(systemId, startDate, endDate, quarter);
        Integer customerMediumPoorEvaluate = customerComplainChartMapper.getCustomerMediumPoorEvaluate(systemId, startDate, endDate, quarter);

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        if(customerComplainQty != null && customerComplainQty != 0){
            entity.setValidComplainRate(numberFormat.format((float)customerValidComplain / customerComplainQty * 100));
            entity.setMediumPoorEvaluateRate(numberFormat.format((float)customerMediumPoorEvaluate / customerComplainQty * 100));
        }

        entity.setCustomerComplain(customerComplainQty);
        entity.setValidComplain(customerValidComplain);
        entity.setMediumPoorEvaluate(customerMediumPoorEvaluate);

        return entity;
    }
}
