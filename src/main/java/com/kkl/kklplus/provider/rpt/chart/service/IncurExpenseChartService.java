package com.kkl.kklplus.provider.rpt.chart.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.mapper.IncurExpenseChartMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class IncurExpenseChartService {

    @Resource
    private IncurExpenseChartMapper incurExpenseChartMapper;

    public List<Double> getIncurExpenseChart(RPTDataDrawingListSearch search) {
        int systemId = RptCommonUtils.getSystemId();
        String yearMonth = DateUtils.getYearMonth(new Date(search.getEndDate()));
        String day = DateUtils.getDay(new Date(search.getEndDate()));
        Integer dayInt = Integer.valueOf(day);

        NameValuePair<BigDecimal, BigDecimal> specialExpenses = incurExpenseChartMapper.getSpecialExpenses(systemId, yearMonth, dayInt);
        double travelCharge = 0.0;
        double praiseCharge = 0.0;
        double timelinessCharge = 0.0;
        double urgentCharge = 0.0;
        double subsidyCharge = 0.0;
        double insuranceCharge = 0.0;
        double otherCharge = 0.0;

        if (specialExpenses != null) {
            travelCharge = specialExpenses.getName().doubleValue();
            otherCharge = specialExpenses.getValue().doubleValue();
        }


        List<Double> list = Lists.newArrayList();
        list.add(travelCharge);
        list.add(praiseCharge);
        list.add(timelinessCharge);
        list.add(urgentCharge);
        list.add(subsidyCharge);
        list.add(insuranceCharge);
        list.add(otherCharge);

        return list;
    }
}
