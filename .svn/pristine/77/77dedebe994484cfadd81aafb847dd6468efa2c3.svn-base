package com.kkl.kklplus.provider.rpt.chart.service;

import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.mapper.KeFuCompleteTimeChartMapper;
import com.kkl.kklplus.provider.rpt.entity.CloseOrderEfficiencyFlagEnum;
import com.kkl.kklplus.provider.rpt.service.KeFuCompleteTimeRptService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuCompleteTimeChartService {

    @Resource
    private KeFuCompleteTimeChartMapper keFuCompleteTimeChartMapper;

    @Autowired
    private KeFuCompleteTimeRptService keFuCompleteTimeRptService;

    /**
     * 获取客服完成时效汇总图表数据
     *
     * @param search
     * @return
     */
    public Map<String, Object> getKeFuCompleteTimeRptData(RPTDataDrawingListSearch search) {

        Date endDate = DateUtils.getEndOfDay(new Date(search.getEndDate()));
        Date startDate = DateUtils.addDays(endDate, -14);
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        String endQuarter = QuarterUtils.getSeasonQuarter(endDate);

        Map<String, Object> map = new HashMap<>();
        if (!quarter.equals(endQuarter)) {
            quarter = null;
        }
        search.setQuarter(quarter);
        search.setBeginDate(startDate.getTime());
        search.setEndDate(endDate.getTime());
        List<RPTKeFuCompleteTimeEntity> list = new ArrayList<>();


        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        List<RPTKeFuCompleteTimeEntity> keFuCompleteTimeData = keFuCompleteTimeChartMapper.getKeFuCompleteTimeInstallData(search);
        if (keFuCompleteTimeData == null || keFuCompleteTimeData.size() <= 0) {
            return map;
        }
        Map<Integer, List<RPTKeFuCompleteTimeEntity>> groupBy = keFuCompleteTimeData.stream().collect(Collectors.groupingBy(RPTKeFuCompleteTimeEntity::getDayIndex));
        int keyInt;
        int sum;
        String key;
        String year;
        Integer month;
        String day;

        int theTotalOrder;
        for (int i = 1; i <= 14; i++) {
            keyInt = StringUtils.toInteger(DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd"));
            sum = 0;
            RPTKeFuCompleteTimeEntity rpt = new RPTKeFuCompleteTimeEntity();
            key = String.valueOf(keyInt);
            year = key.substring(0, 4);
            month = Integer.valueOf(key.substring(4, 6));
            day = key.substring(6, 8);
            if (groupBy.get(keyInt) == null) {
                rpt.setOrderCreateDate(year + "-" + month + "-" + day);
            } else {
                rpt.setOrderCreateDate(year + "-" + month + "-" + day);
                for (RPTKeFuCompleteTimeEntity entity : groupBy.get(keyInt)) {
                    Integer efficiencyFlag = entity.getEfficiencyFlag();
                    if (efficiencyFlag != null) {
                        if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.COMPLETE24HOUR.getValue()) {
                            rpt.setComplete24hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.COMPLETE48HOUR.getValue()) {
                            rpt.setComplete48hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.COMPLETE72HOUR.getValue()) {
                            rpt.setComplete72hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.OVERCOMPLETE72HOUR.getValue()) {
                            rpt.setOverComplete72hour(entity.getEfficiencyFlagSum());
                        }
                    }

                    sum += entity.getEfficiencyFlagSum();
                }
                rpt.setTheTotalOrder(sum);
                theTotalOrder = rpt.getTheTotalOrder();
                keFuCompleteTimeRptService.countCompleteRptRate(numberFormat, rpt, theTotalOrder);

            }
            list.add(rpt);
        }

        list = list.stream().sorted(Comparator.comparing(RPTKeFuCompleteTimeEntity::getOrderCreateDate)).collect(Collectors.toList());


        if (list == null || list.size() <= 0) {
            return map;
        }
        List<String> createDates = new ArrayList<>();
        List<String> strComplete24hourRates = new ArrayList<>();
        List<String> strComplete48hourRates = new ArrayList<>();
        List<String> strComplete72hourRates = new ArrayList<>();
        List<String> strOverComplete72hourRates = new ArrayList<>();
        for (RPTKeFuCompleteTimeEntity entity : list) {
            if (entity.getOrderCreateDate() != null) {
                createDates.add(entity.getOrderCreateDate().substring(5));
                //比率
                strComplete24hourRates.add(entity.getComplete24hourRate());
                strComplete48hourRates.add(entity.getComplete48hourRate());
                strComplete72hourRates.add(entity.getComplete72hourRate());
                strOverComplete72hourRates.add(entity.getOverComplete72hourRate());

            }
        }

        Date date24 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -1));
        Date date48 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -2));
        Date date72 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -3));
        if (endDate.equals(date24)) {

            strComplete48hourRates = strComplete48hourRates.subList(0, 13);
            strComplete72hourRates = strComplete72hourRates.subList(0, 12);
            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 11);
        } else if (endDate.equals(date48)) {

            strComplete72hourRates = strComplete72hourRates.subList(0, 13);
            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 12);
        } else if (endDate.equals(date72)) {

            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 13);
        }
        map.put("createDates", createDates);
        map.put("strComplete24hourRates", strComplete24hourRates);
        map.put("strComplete48hourRates", strComplete48hourRates);
        map.put("strComplete72hourRates", strComplete72hourRates);
        map.put("strOverComplete72hourRates", strOverComplete72hourRates);
        return map;
    }

}
