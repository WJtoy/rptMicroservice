package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.entity.CloseOrderEfficiencyFlagEnum;
import com.kkl.kklplus.provider.rpt.entity.KeFuCompleteTimeRptEntity;
import com.kkl.kklplus.provider.rpt.mapper.KeFuCompleteTimeRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuCompleteTimeRptService extends RptBaseService {

    public static final double MINUTES60 = 60.0;

    @Resource
    private KeFuCompleteTimeRptMapper keFuCompleteTimeRptMapper;

    @Autowired
    private KeFuUtils keFuUtils;

    /**
     * 获取报表显示数据
     *
     * @param search
     * @return
     */
    public List<RPTKeFuCompleteTimeEntity> getKeFuCompleteTimeRptData(RPTKeFuCompleteTimeSearch search) {

        Date endDate = DateUtils.getEndOfDay(new Date(search.getEndDate()));
        Date startDate = DateUtils.addDays(endDate, -31);
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        String endQuarter = QuarterUtils.getSeasonQuarter(endDate);


        if (!quarter.equals(endQuarter)) {
            quarter = null;
        }
        search.setQuarter(quarter);
        search.setBeginDate(startDate.getTime());
        search.setEndDate(endDate.getTime());
        List<RPTKeFuCompleteTimeEntity> list = new ArrayList<>();

        int theTotalOrderSum = 0;
        int complete24hourSum = 0;
        int complete48hourSum = 0;
        int complete72hourSum = 0;
        int cancel24hourSum = 0;
        int cancel48hourSum = 0;
        int cancel72hourSum = 0;
        int overComplete72hourSum = 0;
        int overCancel72hourSum = 0;
        int unfulfilledOrderSum = 0;
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        Set<Long> keFuIds = new HashSet<>();
        Map<Long, RPTUser> kAKeFuMap;
        boolean result = false;
        if(search.getSubFlag() != null && search.getSubFlag() != -1) {
            kAKeFuMap  = MSUserUtils.getMapByUserType(2);
            result = true;
            keFuUtils.getKeFu(kAKeFuMap,search.getSubFlag(),keFuIds);
        }
        List<Long> keFuIdsList =  Lists.newArrayList();
        if(null != search.getKeFuId() && search.getKeFuId() != 0 && result){
            if(keFuIds.size() > 0){
                for(Long id : keFuIds){
                    if(search.getKeFuId().equals(id)){
                        keFuIdsList.add(id);
                    }
                }
                if(keFuIdsList.size() == 0){
                    return list;
                }
            }else {
                return list;
            }
        }else if(null != search.getKeFuId() && search.getKeFuId() != 0){
            keFuIdsList.add(search.getKeFuId());
        }else if(result){
            if(keFuIds.size() > 0){
                keFuIdsList.addAll(Lists.newArrayList(keFuIds));
            }else {
                return list;
            }
        }

        search.setKeFuIds(keFuIdsList);
        List<RPTKeFuCompleteTimeEntity> keFuCompleteTimeData = keFuCompleteTimeRptMapper.getKeFuCompleteTimeData(search);
        if (keFuCompleteTimeData == null || keFuCompleteTimeData.size() <= 0) {
            return list;
        }
        Map<Integer, List<RPTKeFuCompleteTimeEntity>> groupBy = keFuCompleteTimeData.stream().collect(Collectors.groupingBy(RPTKeFuCompleteTimeEntity::getDayIndex));
        int keyInt;
        int sum;
        String key;
        String year;
        String month;
        String day;

        int theTotalOrder;
        for (int i = 1; i <= 31; i++) {
            keyInt = StringUtils.toInteger(DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd"));
            sum = 0;
            RPTKeFuCompleteTimeEntity rpt = new RPTKeFuCompleteTimeEntity();
            key = String.valueOf(keyInt);
            year = key.substring(0, 4);
            month = key.substring(4, 6);
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
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.CANCEL24HOUR.getValue()) {
                            rpt.setCancel24hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.CANCEL48HOUR.getValue()) {
                            rpt.setCancel48hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.CANCEL72HOUR.getValue()) {
                            rpt.setCancel72hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.OVERCOMPLETE72HOUR.getValue()) {
                            rpt.setOverComplete72hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.OVERCANCEL72HOUR.getValue()) {
                            rpt.setOverCancel72hour(entity.getEfficiencyFlagSum());
                        } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.UNFULFILLEDORDER.getValue()) {
                            rpt.setUnfulfilledOrder(entity.getEfficiencyFlagSum());
                        }
                    }

                    sum += entity.getEfficiencyFlagSum();
                }
                rpt.setTheTotalOrder(sum);
                theTotalOrder = rpt.getTheTotalOrder();
                countCompleteRptRate(numberFormat, rpt, theTotalOrder);
                theTotalOrderSum += rpt.getTheTotalOrder();
                complete24hourSum += rpt.getComplete24hour();
                complete48hourSum += rpt.getComplete48hour();
                complete72hourSum += rpt.getComplete72hour();
                cancel24hourSum += rpt.getCancel24hour();
                cancel48hourSum += rpt.getCancel48hour();
                cancel72hourSum += rpt.getCancel72hour();
                overComplete72hourSum += rpt.getOverComplete72hour();
                overCancel72hourSum += rpt.getOverCancel72hour();
                unfulfilledOrderSum += rpt.getUnfulfilledOrder();
            }
            list.add(rpt);
        }

        list = list.stream().sorted(Comparator.comparing(RPTKeFuCompleteTimeEntity::getOrderCreateDate)).collect(Collectors.toList());

        //合计

        RPTKeFuCompleteTimeEntity rptKeFuCompleteTimeEntity = new RPTKeFuCompleteTimeEntity();
        rptKeFuCompleteTimeEntity.setTheTotalOrder(theTotalOrderSum);
        rptKeFuCompleteTimeEntity.setComplete24hour(complete24hourSum);
        rptKeFuCompleteTimeEntity.setComplete48hour(complete48hourSum);
        rptKeFuCompleteTimeEntity.setComplete72hour(complete72hourSum);
        rptKeFuCompleteTimeEntity.setCancel24hour(cancel24hourSum);
        rptKeFuCompleteTimeEntity.setCancel48hour(cancel48hourSum);
        rptKeFuCompleteTimeEntity.setCancel72hour(cancel72hourSum);
        rptKeFuCompleteTimeEntity.setOverComplete72hour(overComplete72hourSum);
        rptKeFuCompleteTimeEntity.setOverCancel72hour(overCancel72hourSum);
        rptKeFuCompleteTimeEntity.setUnfulfilledOrder(unfulfilledOrderSum);
        theTotalOrder = rptKeFuCompleteTimeEntity.getTheTotalOrder();

        countCompleteRptRate(numberFormat, rptKeFuCompleteTimeEntity, theTotalOrder);

        list.add(rptKeFuCompleteTimeEntity);


        return list;

    }

    /**
     * 获取报表数据并转到图表中
     *
     * @return
     */
    public Map<String, Object> turnToChartInformationNew(RPTKeFuCompleteTimeSearch search) {
        Map<String, Object> map = new HashMap<>();
        List<RPTKeFuCompleteTimeEntity> entityList = getKeFuCompleteTimeRptData(search);
        if (entityList == null || entityList.size() <= 0) {
            return map;
        }
        List<String> createDates = new ArrayList<>();
        List<String> strComplete24hours = new ArrayList<>();
        List<String> strComplete48hours = new ArrayList<>();
        List<String> strComplete72hours = new ArrayList<>();
        List<String> strOverComplete72hours = new ArrayList<>();
        List<String> strCancenums = new ArrayList<>();
        List<String> strUnfulfilledOrders = new ArrayList<>();
        List<String> strTheTotalOrders = new ArrayList<>();

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        //比率
        List<String> strComplete24hourRates = new ArrayList<>();
        List<String> strComplete48hourRates = new ArrayList<>();
        List<String> strComplete72hourRates = new ArrayList<>();
        List<String> strOverComplete72hourRates = new ArrayList<>();
        List<String> strCancelNumRates = new ArrayList<>();
        List<String> strUnfulfilledOrderRates = new ArrayList<>();
        Map<String, String> orderCompleteMap = new HashMap<>();
        Map<String, String> complete24Map = new HashMap<>();
        Map<String, String> complete48Map = new HashMap<>();
        Map<String, String> complete72Map = new HashMap<>();
        Map<String, String> overComplete72Map = new HashMap<>();
        Map<String, String> cancelMap = new HashMap<>();
        Map<String, String> unfulfilledOrderMap = new HashMap<>();

        RPTKeFuCompleteTimeEntity orderProcess48hourRptEntity = entityList.get(entityList.size() - 1);
        int orderComplete = orderProcess48hourRptEntity.getComplete24hour() + orderProcess48hourRptEntity.getComplete48hour() +
                orderProcess48hourRptEntity.getComplete72hour() + orderProcess48hourRptEntity.getOverComplete72hour();
        orderCompleteMap.put("value", String.valueOf(orderComplete));
        orderCompleteMap.put("name", "订单完成");
        complete24Map.put("value", orderProcess48hourRptEntity.getComplete24hour().toString());
        complete24Map.put("name", "24小时完成");
        complete48Map.put("value", orderProcess48hourRptEntity.getComplete48hour().toString());
        complete48Map.put("name", "48小时完成");
        complete72Map.put("value", orderProcess48hourRptEntity.getComplete72hour().toString());
        complete72Map.put("name", "72小时完成");
        overComplete72Map.put("value", orderProcess48hourRptEntity.getOverComplete72hour().toString());
        overComplete72Map.put("name", "72小时外完成");
        cancelMap.put("value", orderProcess48hourRptEntity.getCancelTheSum().toString());
        cancelMap.put("name", "取消");
        unfulfilledOrderMap.put("value", orderProcess48hourRptEntity.getUnfulfilledOrder().toString());
        unfulfilledOrderMap.put("name", "未完成");


        List<Map<String, String>> mapList = new ArrayList<>();
        List<Map<String, String>> completeMapList = new ArrayList<>();

        int cancel24hour;
        int cancel48hour;
        int cancel72hour;
        int overCancel72hour;
        int cancelNum;
        for (RPTKeFuCompleteTimeEntity entity : entityList) {
            if (entity.getOrderCreateDate() != null) {
                int theTotalOrder = entity.getTheTotalOrder();

                createDates.add(entity.getOrderCreateDate().substring(5));
                strComplete24hours.add(entity.getComplete24hour().toString());
                strComplete48hours.add(entity.getComplete48hour().toString());
                strComplete72hours.add(entity.getComplete72hour().toString());
                strOverComplete72hours.add(entity.getOverComplete72hour().toString());
                strCancenums.add(entity.getCancelTheSum().toString());
                strUnfulfilledOrders.add(entity.getUnfulfilledOrder().toString());
                strTheTotalOrders.add(entity.getTheTotalOrder().toString());

                if (theTotalOrder != 0) {

                    cancel24hour = entity.getCancel24hour();
                    cancel48hour = entity.getCancel48hour();
                    cancel72hour = entity.getCancel72hour();
                    overCancel72hour = entity.getOverCancel72hour();

                    cancelNum = cancel24hour + cancel48hour + cancel72hour + overCancel72hour;
                    if (cancelNum > 0) {
                        String cancelRate = numberFormat.format((float) cancelNum / theTotalOrder * 100);
                        strCancelNumRates.add(cancelRate);
                    } else {
                        strCancelNumRates.add("0");
                    }

                } else {
                    strCancelNumRates.add("0");
                }
                //比率
                strComplete24hourRates.add(entity.getComplete24hourRate());
                strComplete48hourRates.add(entity.getComplete48hourRate());
                strComplete72hourRates.add(entity.getComplete72hourRate());
                strOverComplete72hourRates.add(entity.getOverComplete72hourRate());
                strUnfulfilledOrderRates.add(entity.getUnfulfilledOrderRate());
            }
        }
        completeMapList.add(orderCompleteMap);
        completeMapList.add(cancelMap);
        completeMapList.add(unfulfilledOrderMap);

        mapList.add(complete24Map);
        mapList.add(complete48Map);
        mapList.add(complete72Map);
        mapList.add(overComplete72Map);
        mapList.add(cancelMap);
        mapList.add(unfulfilledOrderMap);
        map.put("strComplete24hours", strComplete24hours);
        map.put("strComplete48hours", strComplete48hours);
        map.put("strComplete72hours", strComplete72hours);
        map.put("strOverComplete72hours", strOverComplete72hours);
        map.put("strCancenums", strCancenums);
        map.put("strUnfulfilledOrders", strUnfulfilledOrders);
        map.put("strTheTotalOrders", strTheTotalOrders);

        Date endDate = DateUtils.getEndOfDay(new Date(search.getEndDate()));
        Date date24 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -1));
        Date date48 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -2));
        Date date72 = DateUtils.getEndOfDay(DateUtils.addDays(new Date(), -3));

        if (endDate.equals(date24)) {
            strComplete48hourRates = strComplete48hourRates.subList(0, 30);
            strComplete72hourRates = strComplete72hourRates.subList(0, 29);
            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 28);
        } else if (endDate.equals(date48)) {
            strComplete72hourRates = strComplete72hourRates.subList(0, 30);
            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 29);
        } else if (endDate.equals(date72)) {
            strOverComplete72hourRates = strOverComplete72hourRates.subList(0, 30);
        }
        map.put("strComplete24hourRates", strComplete24hourRates);
        map.put("strComplete48hourRates", strComplete48hourRates);
        map.put("strComplete72hourRates", strComplete72hourRates);
        map.put("strOverComplete72hourRates", strOverComplete72hourRates);
        map.put("strCancelNumRates", strCancelNumRates);
        map.put("strUnfulfilledOrderRates", strUnfulfilledOrderRates);

        map.put("mapList", mapList);
        map.put("completeMapList", completeMapList);
        map.put("createDates", createDates);

        return map;
    }

    /**
     * 计算完成报表的比例
     *
     * @param numberFormat
     * @param rpt
     * @param theTotalOrder
     */
    public void countCompleteRptRate(NumberFormat numberFormat, RPTKeFuCompleteTimeEntity rpt, Integer theTotalOrder) {
        if (theTotalOrder != null && theTotalOrder != 0) {
            int complete24hour = rpt.getComplete24hour();
            if (complete24hour > 0) {
                String complete24hourRate = numberFormat.format((float) complete24hour / theTotalOrder * 100);
                rpt.setComplete24hourRate(complete24hourRate);
            }
            int complete48hour = rpt.getComplete48hour() + rpt.getComplete24hour();
            if (complete48hour > 0) {
                String complete48hourRate = numberFormat.format((float) complete48hour / theTotalOrder * 100);
                rpt.setComplete48hourRate(complete48hourRate);
            }
            int complete72hour = rpt.getComplete72hour() + complete48hour;
            if (complete72hour > 0) {
                String complete72hourRate = numberFormat.format((float) complete72hour / theTotalOrder * 100);
                rpt.setComplete72hourRate(complete72hourRate);
            }
            int cancel24hour = rpt.getCancel24hour();
            if (cancel24hour > 0) {
                String cancel24hourRate = numberFormat.format((float) cancel24hour / theTotalOrder * 100);
                rpt.setCancel24hourRate(cancel24hourRate);
            }
            int cancel48hour = rpt.getCancel48hour() + rpt.getCancel24hour();
            if (cancel48hour > 0) {
                String cancel48hourRate = numberFormat.format((float) cancel48hour / theTotalOrder * 100);
                rpt.setCancel48hourRate(cancel48hourRate);
            }

            int cancel72hour = rpt.getCancel72hour() + cancel48hour;
            if (cancel72hour > 0) {
                String cancel72hourRate = numberFormat.format((float) cancel72hour / theTotalOrder * 100);
                rpt.setCancel72hourRate(cancel72hourRate);
            }

            int overComplete72hour = rpt.getOverComplete72hour();
            if (overComplete72hour > 0) {
                String overComplete72hourRate = numberFormat.format((float) overComplete72hour / theTotalOrder * 100);
                rpt.setOverComplete72hourRate(overComplete72hourRate);
            }

            int overCancel72hour = rpt.getOverCancel72hour();
            if (cancel72hour > 0) {
                String overCancel72hourRate = numberFormat.format((float) overCancel72hour / theTotalOrder * 100);
                rpt.setOverCancel72hourRate(overCancel72hourRate);
            }
            int unfulfilledOrder = rpt.getUnfulfilledOrder();
            if (unfulfilledOrder > 0) {
                String unfulfilledOrderRate = numberFormat.format((float) unfulfilledOrder / theTotalOrder * 100);
                rpt.setUnfulfilledOrderRate(unfulfilledOrderRate);
            }
        }
    }

//    public void saveKeFuCompleteTimeToRptDB(Date date) {
//        writeCreateOrderFromYesterday(date);
//        updateOrderInformation(date);
//    }

    /**
     * 删除中间表中指定字段的数据
     */
    private void deleteKeFuCompleteTimeFromRptDB(Date date) {
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int systemId = RptCommonUtils.getSystemId();
            keFuCompleteTimeRptMapper.deleteOrderCreatedData(systemId, beginDate.getTime(), endDate.getTime(), quarter);
        }
    }


    /**
     * 重建下单
     */
    public boolean rebuildMiddleTableCreateOrderData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            writeCreateOrderFromYesterday(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteKeFuCompleteTimeFromRptDB(beginDate);
                            writeCreateOrderFromYesterday(beginDate);
                            break;
                        case DELETE:
                            deleteKeFuCompleteTimeFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTableCreateOrderData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 重建取消/完成
     */
    public boolean rebuildMiddleTableCloseData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            updateOrderCloseData(beginDate);
                            updateOrderCancelledData(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            updateOrderCloseData(beginDate);
                            updateOrderCancelledData(beginDate);
                            break;
                        case DELETE:
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTableCloseData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 重建派单
     */
    public boolean rebuildMiddleTablePlanTypeData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            updateRptPlanType(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            updateRptPlanType(beginDate);
                            break;
                        case DELETE:
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTablePlanTypeData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
    /**
     * 重建投诉
     */
    public boolean rebuildMiddleTableComplainData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            updateRptComplain(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            break;
                        case DELETE:
                            deleteComplainFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTableComplainData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
    /**
     * 写入某天的创建的订单
     *
     * @param date
     */
    public void writeCreateOrderFromYesterday(Date date) {
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Long endDate = DateUtils.getEndOfDay(date).getTime();
        Long startDate = DateUtils.getStartOfDay(date).getTime();


        List<KeFuCompleteTimeRptEntity> list = keFuCompleteTimeRptMapper.getOrderCreatedData(systemId, startDate, endDate, quarter);

        int dayIndex;
        String strDayIndex;
        for (KeFuCompleteTimeRptEntity entity : list) {
            strDayIndex = DateUtils.formatDate(new Date(entity.getOrderCreateDate()), "yyyyMMdd");
            dayIndex = Integer.parseInt(strDayIndex);
            entity.setStatus(10);
            entity.setDayIndex(dayIndex);
        }
        //batch insert
        List<List<KeFuCompleteTimeRptEntity>> parts = Lists.partition(list, 100);
        List<KeFuCompleteTimeRptEntity> part;
        for (List<KeFuCompleteTimeRptEntity> part1 : parts) {
            try {
                part = part1;
                keFuCompleteTimeRptMapper.insertOrderCreatedData(part);
                TimeUnit.MILLISECONDS.sleep(1000);
            } catch (InterruptedException e) {
                log.error("【KeFuCompleteTimeRptService.writeCreateOrderFromYesterday】客服完成时效写入新建订单失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }
        }
    }

//    /**
//     * 按时间更新订单信息
//     *
//     * @param date
//     */
//    public void updateOrderInformation(Date date) {
//
//        int systemId = RptCommonUtils.getSystemId();
//        String quarter = QuarterUtils.getSeasonQuarter(date);
//        Date endDate = DateUtils.getEndOfDay(date);
//        Date startDate = DateUtils.getStartOfDay(date);
//        List<KeFuCompleteTimeRptEntity> closeData = keFuCompleteTimeRptMapper.getOrderCloseData(systemId, startDate.getTime(), endDate.getTime(), quarter);
//        Map<Long, KeFuCompleteTimeRptEntity> closeDataMap = closeData.stream().collect(Collectors.toMap(KeFuCompleteTimeRptEntity::getOrderId, Function.identity(),(key1, key2) -> key2));
//        List<KeFuCompleteTimeRptEntity> cancelledData = keFuCompleteTimeRptMapper.getOrderCancelledData(systemId, startDate.getTime(), endDate.getTime(), quarter);
//        Map<Long, KeFuCompleteTimeRptEntity> cancelledDataMap = cancelledData.stream().collect(Collectors.toMap(KeFuCompleteTimeRptEntity::getOrderId, Function.identity(),(key1, key2) -> key2));
//        List<KeFuCompleteTimeRptEntity> planTypeData = keFuCompleteTimeRptMapper.getPlanTypeData(startDate, endDate);
//        Map<Long, KeFuCompleteTimeRptEntity> planTypeDataMap = planTypeData.stream().collect(Collectors.toMap(KeFuCompleteTimeRptEntity::getOrderId, Function.identity(),(key1, key2) -> key2));
//        List<KeFuCompleteTimeRptEntity> complainData = keFuCompleteTimeRptMapper.getComplainInformation(startDate, endDate);
//        Map<Long, KeFuCompleteTimeRptEntity> complainDataMap = complainData.stream().collect(Collectors.toMap(KeFuCompleteTimeRptEntity::getOrderId, Function.identity(),(key1, key2) -> key2));
//
//        Set<Long> orderIds = new HashSet<>();
//        orderIds.addAll(closeDataMap.keySet());
//        orderIds.addAll(cancelledDataMap.keySet());
//        orderIds.addAll(planTypeDataMap.keySet());
//        orderIds.addAll(complainDataMap.keySet());
//
//        int orderProcessTime;
//        KeFuCompleteTimeRptEntity entity;
//        KeFuCompleteTimeRptEntity rptEntity;
//        try {
//            for (Long key : orderIds) {
//                entity = new KeFuCompleteTimeRptEntity();
//
//                rptEntity = planTypeDataMap.get(key);
//                if (rptEntity != null) {
//                    entity.setOrderId(rptEntity.getOrderId());
//                    entity.setSystemId(systemId);
//                    entity.setServicePointId(rptEntity.getServicePointId());
//                    entity.setPlanType(rptEntity.getPlanType());
//                }
//
//                rptEntity = complainDataMap.get(key);
//                if (rptEntity != null) {
//                    entity.setOrderId(rptEntity.getOrderId());
//                    entity.setSystemId(systemId);
//                    entity.setComplainTime(rptEntity.getComplainTime());
//                }
//
//                rptEntity = closeDataMap.get(key);
//                if (rptEntity != null) {
//
//                    entity.setOrderId(rptEntity.getOrderId());
//                    entity.setSystemId(systemId);
//                    entity.setKeFuId(rptEntity.getKeFuId());
//                    entity.setProvinceId(rptEntity.getProvinceId());
//                    entity.setCityId(rptEntity.getCityId());
//                    entity.setCountyId(rptEntity.getCountyId());
//                    entity.setProductCategoryId(rptEntity.getProductCategoryId());
//                    entity.setServicePointId(rptEntity.getServicePointId());
//                    entity.setOrderCreateDate(rptEntity.getOrderCreateDate());
//                    entity.setOrderCloseDate(rptEntity.getOrderCloseDate());
//                    entity.setStatus(80);
//                    orderProcessTime = getPastMinutes(entity.getOrderCreateDate(), entity.getOrderCloseDate()).intValue();
//                    entity.setOrderProcessTime(orderProcessTime);
//                    countCompleteRptTimeliness(entity, orderProcessTime);
//                }
//
//                rptEntity = cancelledDataMap.get(key);
//                if (rptEntity != null) {
//                    entity.setOrderId(rptEntity.getOrderId());
//                    entity.setSystemId(systemId);
//                    entity.setKeFuId(rptEntity.getKeFuId());
//                    entity.setProvinceId(rptEntity.getProvinceId());
//                    entity.setCityId(rptEntity.getCityId());
//                    entity.setCountyId(rptEntity.getCountyId());
//                    entity.setProductCategoryId(rptEntity.getProductCategoryId());
//                    entity.setServicePointId(rptEntity.getServicePointId());
//                    entity.setOrderCreateDate(rptEntity.getOrderCreateDate());
//                    entity.setOrderCloseDate(rptEntity.getOrderCloseDate());
//                    entity.setStatus(rptEntity.getStatus());
//                    orderProcessTime = getPastMinutes(entity.getOrderCreateDate(), entity.getOrderCloseDate()).intValue();
//                    entity.setOrderProcessTime(orderProcessTime);
//                    countCompleteRptTimeliness(entity, orderProcessTime);
//                }
//
//                keFuCompleteTimeRptMapper.updateOrderInformation(entity);
//            }
//        } catch (Exception e) {
//            log.error("【KeFuCompleteTimeRptService.updateOrderInformation】客服完成时效更新订单信息失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
//        }
//    }

    public void updateOrderCloseData(Date date) {
        Long endDate = DateUtils.getEndOfDay(date).getTime();
        Long startDate = DateUtils.getStartOfDay(date).getTime();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        int systemId = RptCommonUtils.getSystemId();
        List<KeFuCompleteTimeRptEntity> list = keFuCompleteTimeRptMapper.getOrderCloseData(systemId, startDate, endDate,quarter);
        int orderProcessTime;
        try {
            for (KeFuCompleteTimeRptEntity entity : list) {
                entity.setStatus(80);
                if(entity.getOrderArrivalDate() != null && entity.getOrderArrivalDate() != 0 ){
                    orderProcessTime = getPastMinutes(entity.getOrderArrivalDate(), entity.getOrderCloseDate()).intValue();
                }else{
                    orderProcessTime = getPastMinutes(entity.getOrderCreateDate(), entity.getOrderCloseDate()).intValue();
                }
                entity.setOrderProcessTime(orderProcessTime);
                countCompleteRptTimeliness(entity, orderProcessTime);
                keFuCompleteTimeRptMapper.updateOrderCloseData(entity);
            }
        } catch (Exception e) {
            log.error("【KeFuCompleteTimeRptService.updateOrderCloseData】客服完成时效更新完成订单失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }
    }

    public void updateOrderCancelledData(Date date) {
        Long endDate = DateUtils.getEndOfDay(date).getTime();
        Long startDate = DateUtils.getStartOfDay(date).getTime();
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<KeFuCompleteTimeRptEntity> list = keFuCompleteTimeRptMapper.getOrderCancelledData(systemId, startDate, endDate,quarter);
        int orderProcessTime;
        try {
            for (KeFuCompleteTimeRptEntity entity : list) {
                orderProcessTime = getPastMinutes(entity.getOrderCreateDate(), entity.getOrderCloseDate()).intValue();
                entity.setOrderProcessTime(orderProcessTime);
                countCompleteRptTimeliness(entity, orderProcessTime);
                keFuCompleteTimeRptMapper.updateOrderCancelledData(entity);
            }
        } catch (Exception e) {
            log.error("【KeFuCompleteTimeRptService.updateOrderCancelledData】客服完成时效更新取消信息失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }
    }



    //更新派单类型
    public void updateRptPlanType(Date date) {
        Date startDate = DateUtils.getStartOfDay(date);
        Date endDate = DateUtils.getEndOfDay(date);
        int systemId = RptCommonUtils.getSystemId();
        List<KeFuCompleteTimeRptEntity> updateDataAllList = Lists.newArrayList();
        List<KeFuCompleteTimeRptEntity> updateData;
        int endLimit = 20000;
        int beginLimit;
        int pageNo = 1;
        int listSize = 0;
        do{
            beginLimit = (pageNo - 1) * endLimit;
            updateData = keFuCompleteTimeRptMapper.getPlanTypeData(startDate, endDate,beginLimit,endLimit);
            if(updateData != null){
                updateDataAllList.addAll(updateData);
                listSize = updateData.size();
            }
            pageNo++;
        }while (listSize == endLimit);

        if (!updateDataAllList.isEmpty()) {
            for (KeFuCompleteTimeRptEntity entity : updateDataAllList) {
                entity.setSystemId(systemId);
                keFuCompleteTimeRptMapper.updateRPtPlanType(entity);
            }
        }
    }

    //更新投诉信息
    public void updateRptComplain(Date date) {
        Date startDate = DateUtils.getStartOfDay(date);
        Date endDate = DateUtils.getEndOfDay(date);
        int systemId = RptCommonUtils.getSystemId();
        List<KeFuCompleteTimeRptEntity> list = keFuCompleteTimeRptMapper.getComplainInformation(startDate, endDate);
        for (KeFuCompleteTimeRptEntity entity : list) {
            entity.setSystemId(systemId);
            keFuCompleteTimeRptMapper.updateComplainInformation(entity);
        }
    }

    /**
     * 删除中间表中指定字段的数据
     */
    private void deleteComplainFromRptDB(Date date) {
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int systemId = RptCommonUtils.getSystemId();
            keFuCompleteTimeRptMapper.deleteComplainData(systemId, beginDate.getTime(), endDate.getTime(), quarter);
        }
    }

    private Long getPastMinutes(Long startDate, Long endDate){
        if(startDate > endDate){
            return 1L;
        }
        long t =  (endDate - startDate);
        long minute = t / (60 * 1000);
        if(minute == 0){
            minute = 1;
        }
        return minute;
    }
    /**
     * 计算完成报表的完成时效
     *
     * @param entity
     * @param orderProcessTime
     */
    private void countCompleteRptTimeliness(KeFuCompleteTimeRptEntity entity, int orderProcessTime) {
        if (orderProcessTime > 0) {
            if (entity.getStatus() == 80 || entity.getStatus() == 85) {
                if (orderProcessTime / MINUTES60 <= 24) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.COMPLETE24HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 <= 48) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.COMPLETE48HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 <= 72) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.COMPLETE72HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 > 72) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.OVERCOMPLETE72HOUR.getValue());
                }
            } else if (entity.getStatus() > 85) {
                if (orderProcessTime / MINUTES60 <= 24) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.CANCEL24HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 <= 48) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.CANCEL48HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 <= 72) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.CANCEL72HOUR.getValue());
                } else if (orderProcessTime / MINUTES60 > 72) {
                    entity.setEfficiencyFlag(CloseOrderEfficiencyFlagEnum.OVERCANCEL72HOUR.getValue());
                }
            }
        } else {
            entity.setEfficiencyFlag(0);
        }
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = false;
        RPTKeFuCompleteTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuCompleteTimeSearch.class);
        Date endDate = DateUtils.getEndOfDay(new Date(searchCondition.getEndDate()));
        Date startDate = DateUtils.addDays(endDate, -31);
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        String endQuarter = QuarterUtils.getSeasonQuarter(endDate);


        if (!quarter.equals(endQuarter)) {
            quarter = null;
        }
        searchCondition.setQuarter(quarter);
        searchCondition.setBeginDate(startDate.getTime());
        searchCondition.setEndDate(endDate.getTime());
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getEndDate() != null) {
            Set<Long> keFuIds = new HashSet<>();
            Map<Long, RPTUser> kAKeFuMap;
            if(searchCondition.getSubFlag() != null && searchCondition.getSubFlag() != -1) {
                result = true;
                kAKeFuMap = MSUserUtils.getMapByUserType(2);
                keFuUtils.getKeFu(kAKeFuMap,searchCondition.getSubFlag(),keFuIds);
            }
            List<Long> keFuIdsList =  Lists.newArrayList();
            List<Long> keFuSubTypeList =  Lists.newArrayList(keFuIds);
            if(searchCondition.getKeFuId() != 0 && result){
                if(keFuSubTypeList.size() > 0){
                    for(Long id : keFuSubTypeList){
                        if(searchCondition.getKeFuId().equals(id)){
                            keFuIdsList.add(id);
                        }
                    }
                    if(keFuIdsList.size() == 0){
                        return false;
                    }
                }else {
                    return false;
                }
            }else if(searchCondition.getKeFuId() != 0){
                keFuIdsList.add(searchCondition.getKeFuId());
            }else if(result){
                if(keFuSubTypeList.size() > 0){
                    keFuIdsList.addAll(keFuSubTypeList);
                }else {
                    return false;
                }
            }
            searchCondition.setKeFuIds(keFuIdsList);
            Integer rowCount = keFuCompleteTimeRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook keFuCompleteTimeRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTKeFuCompleteTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuCompleteTimeSearch.class);
            List<RPTKeFuCompleteTimeEntity> list = getKeFuCompleteTimeRptData(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //===============绘制标题======================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 19));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单日期");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 0, 0));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单数量");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 1, 1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "24小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 2, 5));
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "48小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 6, 9));
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "72小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 10, 13));
            ExportExcel.createCell(headFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "72小时外");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 14, 17));
            ExportExcel.createCell(headFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完成");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 18, 18));
            ExportExcel.createCell(headFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 19, 19));

            //表头第二行===============================
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuCompleteTimeEntity orderProcess48hourRptEntity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    if (dataRowIndex == list.size() - 1) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                    } else {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOrderCreateDate());
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getTheTotalOrder());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete24hour());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete24hourRate() + "%");
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel24hour());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel24hourRate() + "%");
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete48hour() + orderProcess48hourRptEntity.getComplete24hour());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete48hourRate() + "%");
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel48hour() + orderProcess48hourRptEntity.getCancel24hour());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel48hourRate() + "%");
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete72hour() +
                            orderProcess48hourRptEntity.getComplete48hour() + orderProcess48hourRptEntity.getComplete24hour());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete72hourRate() + "%");
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel72hour() +
                            orderProcess48hourRptEntity.getCancel48hour() + orderProcess48hourRptEntity.getCancel24hour());
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getCancel72hourRate() + "%");

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverComplete72hour());
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverComplete72hourRate() + "%");
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverCancel72hour());
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverCancel72hourRate() + "%");
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getUnfulfilledOrder());
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getUnfulfilledOrderRate() + "%");
                }
            }
        } catch (Exception e) {
            log.error("【KeFuCompleteTimeRptService.keFuCompleteTimeRptExport】客服完工时效报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
