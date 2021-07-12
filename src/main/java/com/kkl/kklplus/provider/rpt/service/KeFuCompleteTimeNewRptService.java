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
public class KeFuCompleteTimeNewRptService extends RptBaseService {

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
        int overComplete72hourSum = 0;
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
        List<RPTKeFuCompleteTimeEntity> keFuCompleteTimeData = keFuCompleteTimeRptMapper.getKeFuCompleteTimeDataNew(search);
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
                       } else if (efficiencyFlag == CloseOrderEfficiencyFlagEnum.OVERCOMPLETE72HOUR.getValue()) {
                            rpt.setOverComplete72hour(entity.getEfficiencyFlagSum());
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
                overComplete72hourSum += rpt.getOverComplete72hour();
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
        rptKeFuCompleteTimeEntity.setOverComplete72hour(overComplete72hourSum);
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
        List<String> strUnfulfilledOrders = new ArrayList<>();
        List<String> strTheTotalOrders = new ArrayList<>();

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        //比率
        List<String> strComplete24hourRates = new ArrayList<>();
        List<String> strComplete48hourRates = new ArrayList<>();
        List<String> strComplete72hourRates = new ArrayList<>();
        List<String> strOverComplete72hourRates = new ArrayList<>();
        List<String> strUnfulfilledOrderRates = new ArrayList<>();
        Map<String, String> orderCompleteMap = new HashMap<>();
        Map<String, String> complete24Map = new HashMap<>();
        Map<String, String> complete48Map = new HashMap<>();
        Map<String, String> complete72Map = new HashMap<>();
        Map<String, String> overComplete72Map = new HashMap<>();
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
        unfulfilledOrderMap.put("value", orderProcess48hourRptEntity.getUnfulfilledOrder().toString());
        unfulfilledOrderMap.put("name", "未完成");


        List<Map<String, String>> mapList = new ArrayList<>();
        List<Map<String, String>> completeMapList = new ArrayList<>();


        for (RPTKeFuCompleteTimeEntity entity : entityList) {
            if (entity.getOrderCreateDate() != null) {

                createDates.add(entity.getOrderCreateDate().substring(5));
                strComplete24hours.add(entity.getComplete24hour().toString());
                strComplete48hours.add(entity.getComplete48hour().toString());
                strComplete72hours.add(entity.getComplete72hour().toString());
                strOverComplete72hours.add(entity.getOverComplete72hour().toString());
                strUnfulfilledOrders.add(entity.getUnfulfilledOrder().toString());
                strTheTotalOrders.add(entity.getTheTotalOrder().toString());

                //比率
                strComplete24hourRates.add(entity.getComplete24hourRate());
                strComplete48hourRates.add(entity.getComplete48hourRate());
                strComplete72hourRates.add(entity.getComplete72hourRate());
                strOverComplete72hourRates.add(entity.getOverComplete72hourRate());
                strUnfulfilledOrderRates.add(entity.getUnfulfilledOrderRate());
            }
        }
        completeMapList.add(orderCompleteMap);
        completeMapList.add(unfulfilledOrderMap);

        mapList.add(complete24Map);
        mapList.add(complete48Map);
        mapList.add(complete72Map);
        mapList.add(overComplete72Map);
        mapList.add(unfulfilledOrderMap);
        map.put("strComplete24hours", strComplete24hours);
        map.put("strComplete48hours", strComplete48hours);
        map.put("strComplete72hours", strComplete72hours);
        map.put("strOverComplete72hours", strOverComplete72hours);
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

            int overComplete72hour = rpt.getOverComplete72hour();
            if (overComplete72hour > 0) {
                String overComplete72hourRate = numberFormat.format((float) overComplete72hour / theTotalOrder * 100);
                rpt.setOverComplete72hourRate(overComplete72hourRate);
            }
            int unfulfilledOrder = rpt.getUnfulfilledOrder();
            if (unfulfilledOrder > 0) {
                String unfulfilledOrderRate = numberFormat.format((float) unfulfilledOrder / theTotalOrder * 100);
                rpt.setUnfulfilledOrderRate(unfulfilledOrderRate);
            }
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
            Integer rowCount = keFuCompleteTimeRptMapper.hasReportNewData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook keFuCompleteTimeRptNewExport(String searchConditionJson, String reportTitle) {
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单日期");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 0, 0));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单数量");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 1, 1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "24小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 2, 3));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "48小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 4, 5));
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "72小时");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 6, 7));
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "72小时外");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 8, 9));
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完成");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 10, 10));
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 11, 11));

            //表头第二行===============================
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "比率");
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

                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete48hour() + orderProcess48hourRptEntity.getComplete24hour());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete48hourRate() + "%");

                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete72hour() +
                            orderProcess48hourRptEntity.getComplete48hour() + orderProcess48hourRptEntity.getComplete24hour());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getComplete72hourRate() + "%");


                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverComplete72hour());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getOverComplete72hourRate() + "%");
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getUnfulfilledOrder());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderProcess48hourRptEntity.getUnfulfilledOrderRate() + "%");
                }
            }
        } catch (Exception e) {
            log.error("【KeFuCompleteTimeRptService.keFuCompleteTimeRptNewExport】客服完工时效报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
