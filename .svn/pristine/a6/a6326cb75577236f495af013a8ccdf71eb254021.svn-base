package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderTimeEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.ChargeBaseEntity;
import com.kkl.kklplus.provider.rpt.entity.CustomerOrderTimeRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CustomerOrderTimeRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
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
public class CustomerOrderTimeRptService extends RptBaseService{
    @Resource
    private CustomerOrderTimeRptMapper customerOrderTimeRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;

    /**
     * 从新表当中查询数据用于报表展示
     *
     */
    public List<RPTCustomerOrderTimeEntity> getCustomerOrderTimeRptData(RPTCustomerOrderTimeSearch search) {

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
        List<RPTCustomerOrderTimeEntity> list = Lists.newArrayList();
        List<ChargeBaseEntity> recordCount = customerOrderTimeRptMapper.getRecordCount(search);

        if (recordCount.isEmpty()) {
            return list;
        }


        //派单时效
        List<ChargeBaseEntity> lessSix = customerOrderTimeRptMapper.getLessSix(search);
        Map<Integer, Long> lessSixMap = lessSix.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));

        List<ChargeBaseEntity> moreSix = customerOrderTimeRptMapper.getMoreSix(search);
        Map<Integer, Long> moreSixMap = moreSix.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));


        //接单时效
        List<ChargeBaseEntity> close12to24 = customerOrderTimeRptMapper.getClose12to24(search);
        Map<Integer, Long> close12to24Map = close12to24.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));

        List<ChargeBaseEntity> closeLessTwelve = customerOrderTimeRptMapper.getCloseLessTwelve(search);
        Map<Integer, Long> closeLessTwelveMap = closeLessTwelve.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));

        List<ChargeBaseEntity> close24to48 = customerOrderTimeRptMapper.getClose24to48(search);
        Map<Integer, Long> close24to48Map = close24to48.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));

        List<ChargeBaseEntity> close48to72 = customerOrderTimeRptMapper.getClose48to72(search);
        Map<Integer, Long> close48to72Map = close48to72.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));

        List<ChargeBaseEntity> closeMore72 = customerOrderTimeRptMapper.getCloseMore72(search);
        Map<Integer, Long> closeMore72Map = closeMore72.stream().collect(Collectors.toMap(ChargeBaseEntity::getDayIndex, ChargeBaseEntity::getCountNum, (key1, key2) -> key2));


        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        long lessSixSum = 0;
        long moreSixSum = 0;
        int lessTwelveSum = 0;
        int close12to24Sum = 0;
        int close24to48Sum = 0;
        int close48to72Sum = 0;
        int closeMore72Sum = 0;
        long lessSixNum;
        long moreSixNum;
        int closeSum = 0;

        long close;
        Long lessTwelve;
        Long less24Hour;
        Long less48Hour;
        Long less72Hour;
        Long more72Hour;


        String key;
        String year;
        String month;
        String day;
        Integer keyInt;
        String completeLessTwelvePercentage;
        String complete12to24Percentage;
        String complete24to48Percentage;
        String complete48to72Percentage;
        String completeMore72Percentage;
        String less24Percentage;
        String less48Percentage;
        String less72Percentage;

        for (int i = 1; i <= 31; i++) {
            keyInt = StringUtils.toInteger(DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd"));
            key = String.valueOf(keyInt);
            RPTCustomerOrderTimeEntity customerOrderTimeEntity = new RPTCustomerOrderTimeEntity();

            year = key.substring(0, 4);
            month = key.substring(4, 6);
            day = key.substring(6, 8);

            lessSixNum = 0;
            moreSixNum = 0;

            lessTwelve = 0L;
            less24Hour = 0L;
            less48Hour = 0L;
            less72Hour = 0L;
            more72Hour = 0L;

            completeLessTwelvePercentage = "0";
            complete12to24Percentage = "0";
            complete24to48Percentage = "0";
            complete48to72Percentage = "0";
            completeMore72Percentage = "0";
            less24Percentage = "0";
            less48Percentage = "0";
            less72Percentage = "0";

            customerOrderTimeEntity.setOrderCreateDate(year + "-" + month + "-" + day);
            if (lessSixMap.get(keyInt) != null) {
                lessSixNum = lessSixMap.get(keyInt);
            }
            if (moreSixMap.get(keyInt) != null) {
                moreSixNum = moreSixMap.get(keyInt);
            }
            if (closeLessTwelveMap.get(keyInt) != null) {
                lessTwelve = closeLessTwelveMap.get(keyInt);
            }
            if (close12to24Map.get(keyInt) != null) {
                less24Hour = close12to24Map.get(keyInt);
            }
            if (close24to48Map.get(keyInt) != null) {
                less48Hour = close24to48Map.get(keyInt);
            }
            if(close48to72Map.get(keyInt) != null){
                less72Hour = close48to72Map.get(keyInt);
            }
            if(closeMore72Map.get(keyInt) != null){
                more72Hour = closeMore72Map.get(keyInt);
            }

            customerOrderTimeEntity.setPlanTimeLessSix(String.valueOf(lessSixNum));

            customerOrderTimeEntity.setPlanTimeMoreSix(String.valueOf(moreSixNum));

            customerOrderTimeEntity.setCloseLessTwelve(lessTwelve.intValue());
            customerOrderTimeEntity.setCloseTwelveBetweenOneDay(less24Hour.intValue());
            customerOrderTimeEntity.setCloseOneDayBetweenTwoDay(less48Hour.intValue());
            customerOrderTimeEntity.setCloseTwoDayBetweenThreeDay(less72Hour.intValue());
            customerOrderTimeEntity.setCloseMoreThreeDay(more72Hour.intValue());
            customerOrderTimeEntity.setCloseSum(customerOrderTimeEntity.getCloseLessTwelve() + customerOrderTimeEntity.getCloseTwelveBetweenOneDay() +
                    customerOrderTimeEntity.getCloseOneDayBetweenTwoDay() + customerOrderTimeEntity.getCloseTwoDayBetweenThreeDay() + customerOrderTimeEntity.getCloseMoreThreeDay());

            close = customerOrderTimeEntity.getCloseSum();

            if (close != 0) {
                completeLessTwelvePercentage = numberFormat.format((float) lessTwelve / close * 100);
                complete12to24Percentage = numberFormat.format((float) less24Hour / close * 100);
                complete24to48Percentage = numberFormat.format((float) less48Hour / close * 100);
                complete48to72Percentage = numberFormat.format((float) less72Hour / close * 100);
                completeMore72Percentage = numberFormat.format((float) more72Hour / close * 100);
                less24Percentage = numberFormat.format((float) (less24Hour + lessTwelve) / close * 100);
                less48Percentage = numberFormat.format((float) (less48Hour + lessTwelve + less24Hour) / close * 100);
                less72Percentage = numberFormat.format((float) (less72Hour + less48Hour + lessTwelve + less24Hour) / close * 100);
            }
            customerOrderTimeEntity.setLessTwelveProportion(completeLessTwelvePercentage);
            customerOrderTimeEntity.setTwelveBetweenOneDayProportion(complete12to24Percentage);
            customerOrderTimeEntity.setOneDayBetweenTwoDayProportion(complete24to48Percentage);
            customerOrderTimeEntity.setTwoDayBetweenThreeDayProportion(complete48to72Percentage);
            customerOrderTimeEntity.setLess24Proportion(less24Percentage);
            customerOrderTimeEntity.setLess48Proportion(less48Percentage);
            customerOrderTimeEntity.setLess72Proportion(less72Percentage);
            customerOrderTimeEntity.setMoreThreeDayProportion(completeMore72Percentage);

            lessSixSum += lessSixNum;
            moreSixSum += moreSixNum;
            lessTwelveSum += customerOrderTimeEntity.getCloseLessTwelve();
            close12to24Sum += customerOrderTimeEntity.getCloseTwelveBetweenOneDay();
            close24to48Sum += customerOrderTimeEntity.getCloseOneDayBetweenTwoDay();
            close48to72Sum += customerOrderTimeEntity.getCloseTwoDayBetweenThreeDay();
            closeMore72Sum += customerOrderTimeEntity.getCloseMoreThreeDay();
            closeSum += customerOrderTimeEntity.getCloseSum();
            list.add(customerOrderTimeEntity);
        }
        //合计

        RPTCustomerOrderTimeEntity customerOrderTimeEntity = new RPTCustomerOrderTimeEntity();
        customerOrderTimeEntity.setCloseSum(closeSum);
        customerOrderTimeEntity.setPlanTimeLessSix(String.valueOf(lessSixSum));
        customerOrderTimeEntity.setPlanTimeMoreSix(String.valueOf(moreSixSum));

        customerOrderTimeEntity.setCloseMoreThreeDay(closeMore72Sum);
        customerOrderTimeEntity.setCloseTwoDayBetweenThreeDay(close48to72Sum + close24to48Sum + close12to24Sum + lessTwelveSum);
        customerOrderTimeEntity.setCloseOneDayBetweenTwoDay(close24to48Sum + close12to24Sum + lessTwelveSum);
        customerOrderTimeEntity.setCloseTwelveBetweenOneDay(close12to24Sum + lessTwelveSum);
        customerOrderTimeEntity.setCloseLessTwelve(lessTwelveSum);

        //计算比例

        String closeLessTwelvePercentage = numberFormat.format((float) lessTwelveSum / closeSum * 100);
        String close12to24Percentage = numberFormat.format((float) (close12to24Sum + lessTwelveSum) / closeSum * 100);
        String close24to48Percentage = numberFormat.format((float) (close24to48Sum + close12to24Sum + lessTwelveSum) / closeSum * 100);
        String close48to72Percentage = numberFormat.format((float) (close48to72Sum + close24to48Sum + close12to24Sum + lessTwelveSum) / closeSum * 100);
        String closeMore72Percentage = numberFormat.format((float) closeMore72Sum / closeSum * 100);
        customerOrderTimeEntity.setLessTwelveProportion(closeLessTwelvePercentage);
        customerOrderTimeEntity.setTwelveBetweenOneDayProportion(close12to24Percentage);
        customerOrderTimeEntity.setOneDayBetweenTwoDayProportion(close24to48Percentage);
        customerOrderTimeEntity.setTwoDayBetweenThreeDayProportion(close48to72Percentage);
        customerOrderTimeEntity.setMoreThreeDayProportion(closeMore72Percentage);
        list.add(customerOrderTimeEntity);

        return list;
    }

    public Map<String, Object> getCustomerOrderTimeChart(RPTCustomerOrderTimeSearch search) {
        Map<String, Object> map = new HashMap<>();
        List<RPTCustomerOrderTimeEntity> processTimeRpt = getCustomerOrderTimeRptData(search);
        if (processTimeRpt == null || processTimeRpt.size() == 0) {
            return map;
        }
        //派单饼状图 数据
        Map<String, String> less3PlanMap = new HashMap<>();
        Map<String, String> more3PlanMap = new HashMap<>();

        RPTCustomerOrderTimeEntity customerOrderTimeEntity = processTimeRpt.get(processTimeRpt.size() - 1);
        less3PlanMap.put("name", "小于3小时");
        less3PlanMap.put("value", customerOrderTimeEntity.getPlanTimeLessSix());
        more3PlanMap.put("name", "大于3小时");
        more3PlanMap.put("value", customerOrderTimeEntity.getPlanTimeMoreSix());
        //完成时效饼状图数据
        Map<String, String> less12HourMap = new HashMap<>();
        Map<String, String> less24HourMap = new HashMap<>();
        Map<String, String> less48HourMap = new HashMap<>();
        Map<String, String> less72HourMap = new HashMap<>();
        Map<String, String> more72HourMap = new HashMap<>();
        less12HourMap.put("name", "小于12小时");
        less12HourMap.put("value", String.valueOf(customerOrderTimeEntity.getCloseLessTwelve()));
        less24HourMap.put("name", "小于24小时");
        less24HourMap.put("value", String.valueOf(customerOrderTimeEntity.getCloseTwelveBetweenOneDay()));
        less48HourMap.put("name", "小于48小时");
        less48HourMap.put("value", String.valueOf(customerOrderTimeEntity.getCloseOneDayBetweenTwoDay()));
        less72HourMap.put("name", "小于72小时");
        less72HourMap.put("value", String.valueOf(customerOrderTimeEntity.getCloseTwoDayBetweenThreeDay()));
        more72HourMap.put("name", "大于72小时");
        more72HourMap.put("value", String.valueOf(customerOrderTimeEntity.getCloseMoreThreeDay()));

        List<Map<String, String>> mapPlanList = Lists.newArrayList();
        mapPlanList.add(less3PlanMap);
        mapPlanList.add(more3PlanMap);

        List<Map<String, String>> mapCloseList = Lists.newArrayList();
        mapCloseList.add(less12HourMap);
        mapCloseList.add(less24HourMap);
        mapCloseList.add(less48HourMap);
        mapCloseList.add(less72HourMap);
        mapCloseList.add(more72HourMap);
        //柱状图
        List<String> less12List = Lists.newArrayList();
        List<String> less24List = Lists.newArrayList();
        List<String> less48List = Lists.newArrayList();
        List<String> less72List = Lists.newArrayList();
        List<String> more72List = Lists.newArrayList();

        //折线图  比例
        List<String> less12RateList = Lists.newArrayList();
        List<String> less24RateList = Lists.newArrayList();
        List<String> less48RateList = Lists.newArrayList();
        List<String> less72RateList = Lists.newArrayList();
        List<String> more72RateList = Lists.newArrayList();
        List<String> dateList = Lists.newArrayList();
        List<String> dateAllList = Lists.newArrayList();


        for (RPTCustomerOrderTimeEntity entity : processTimeRpt) {
            if (entity.getOrderCreateDate() != null) {
                dateAllList.add(entity.getOrderCreateDate().substring(5));
                dateList.add(entity.getOrderCreateDate().substring(5));

                //数量
                less12List.add(String.valueOf(entity.getCloseLessTwelve()));
                less24List.add(String.valueOf(entity.getCloseTwelveBetweenOneDay()));
                less48List.add(String.valueOf(entity.getCloseOneDayBetweenTwoDay()));
                less72List.add(String.valueOf(entity.getCloseTwoDayBetweenThreeDay()));
                more72List.add(String.valueOf(entity.getCloseMoreThreeDay()));
                //比例
                less12RateList.add(entity.getLessTwelveProportion() == null ? "0" : entity.getLessTwelveProportion());
                less24RateList.add(entity.getLess24Proportion() == null ? "0" : entity.getLess24Proportion());
                less48RateList.add(entity.getLess48Proportion() == null ? "0" : entity.getLess48Proportion());
                less72RateList.add(entity.getLess72Proportion() == null ? "0" : entity.getLess72Proportion());
                more72RateList.add(entity.getMoreThreeDayProportion() == null ? "0" : entity.getMoreThreeDayProportion());
            }

        }

        map.put("mapPlanList", mapPlanList);
        map.put("mapCloseList", mapCloseList);
        map.put("dateList", dateList);

        map.put("dateAllList", dateAllList);
        map.put("less12List", less12List);
        map.put("less24List", less24List);
        map.put("less48List", less48List);
        map.put("less72List", less72List);
        map.put("more72List", more72List);

        map.put("less12RateList", less12RateList);
        map.put("less24RateList", less24RateList);
        map.put("less48RateList", less48RateList);
        map.put("less72RateList", less72RateList);
        map.put("more72RateList", more72RateList);

        return map;
    }
    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderTimeSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = customerOrderTimeRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook customerOrderTimeRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderTimeSearch.class);
            List<RPTCustomerOrderTimeEntity> processTimeRpt = getCustomerOrderTimeRptData(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);


            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "天/月");

            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单时效");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 1, 2));

            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结单时效");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 3, 12));

            ExportExcel.createCell(headFirstRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成合计(单)");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 13, 13));


            //表头第二行===============================
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于3小时");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "大于3小时");

            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于12小时");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于12小时完成率");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于24小时");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于24小时完成率");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于48小时");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于48小时完成率");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于72小时");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "小于72小时完成率");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "大于72小时");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "大于72小时完成率");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            if (processTimeRpt != null && processTimeRpt.size() > 0) {
                int rowsCount = processTimeRpt.size();

                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    if (dataRowIndex < 31) {
                        RPTCustomerOrderTimeEntity processTimeListEntity = processTimeRpt.get(dataRowIndex);
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getOrderCreateDate());
                        if (processTimeListEntity.getCloseSum() != 0) {
//
                            ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getPlanTimeLessSix());
                            ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getPlanTimeMoreSix());
                            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseLessTwelve());
                            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getLessTwelveProportion() + "%");
                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseTwelveBetweenOneDay() + processTimeListEntity.getCloseLessTwelve());
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getLess24Proportion() + "%");
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseOneDayBetweenTwoDay() + processTimeListEntity.getCloseTwelveBetweenOneDay() + processTimeListEntity.getCloseLessTwelve());
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getLess48Proportion() + "%");
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseTwoDayBetweenThreeDay() + processTimeListEntity.getCloseOneDayBetweenTwoDay() + processTimeListEntity.getCloseTwelveBetweenOneDay() + processTimeListEntity.getCloseLessTwelve());
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getLess72Proportion() + "%");
                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseMoreThreeDay());
                            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getMoreThreeDayProportion() + "%");
                            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseSum());
                        }
                    }
                    if (dataRowIndex == rowsCount - 1) {
                        RPTCustomerOrderTimeEntity processTimeListEntity = processTimeRpt.get(dataRowIndex);
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getPlanTimeLessSix());
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getPlanTimeMoreSix());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseLessTwelve());
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getLessTwelveProportion() + "%");
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseTwelveBetweenOneDay());
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getTwelveBetweenOneDayProportion() + "%");
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseOneDayBetweenTwoDay());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getOneDayBetweenTwoDayProportion() + "%");
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseTwoDayBetweenThreeDay());
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getTwoDayBetweenThreeDayProportion() + "%");
                        ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseMoreThreeDay());
                        ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getMoreThreeDayProportion() + "%");
                        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, processTimeListEntity.getCloseSum());
                    }

                }
            }

        } catch (Exception e) {
            log.error("【CustomerOrderTimeRptService.customerOrderTimeRptExport】客户工单时效报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
    /**
     * 根据时间获取原始数据
     * @param date
     * @return
     */
    private List<CustomerOrderTimeRptEntity> getCustomerOrderTimeData(Date date){
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date startDate = DateUtils.getStartOfDay(date);
        Date endDate = DateUtils.getEndOfDay(date);
        List<CustomerOrderTimeRptEntity> list = new ArrayList<>();
        List<CustomerOrderTimeRptEntity> customerOrderTimeList = customerOrderTimeRptMapper.getCustomerOrderTimeData(startDate, endDate);
        int dayIndex;
        String strDayIndex;
        Map<Long,RPTArea> areaMap = areaCacheService.getAllCountyMap();
        for(CustomerOrderTimeRptEntity entity : customerOrderTimeList){
            entity.setSystemId(systemId);
            if (entity.getFirstPlanDate() == null && entity.getPlanDate() != null) {
                entity.setFirstPlanDate(entity.getPlanDate());
            }
            entity.setOrderCreateDt(entity.getOrderCreateDate().getTime());
            entity.setOrderCloseDt(entity.getOrderCloseDate().getTime());
            entity.setFirstPlanDt(entity.getFirstPlanDate().getTime());
            RPTArea area = areaMap.get(entity.getCountyId());
            if(area != null){
                String[] split = area.getParentIds().split(",");
                if(split.length == 4){
                    entity.setCityId(Long.valueOf(split[3]));
                    entity.setProvinceId(Long.valueOf(split[2]));
                }
            }
            entity.setCreatePlanTime((entity.getFirstPlanDt().intValue() - entity.getOrderCreateDt().intValue()) / 1000);
            entity.setPlanCloseTime((entity.getOrderCloseDt().intValue() - entity.getFirstPlanDt().intValue()) / 1000);
            strDayIndex = DateUtils.formatDate(entity.getOrderCloseDate(), "yyyyMMdd");
            dayIndex = Integer.parseInt(strDayIndex);
            entity.setQuarter(quarter);
            entity.setDayIndex(dayIndex);
            list.add(entity);
        }
        return list;
    }


    public void saveCustomerOrderTimeToRptDB(Date date){
        List<CustomerOrderTimeRptEntity> list = getCustomerOrderTimeData(date);

        List<List<CustomerOrderTimeRptEntity>> parts = Lists.partition(list, 100);
        List<CustomerOrderTimeRptEntity> part;
        for (int i = 0, size = parts.size(); i < size; i++) {
            part = parts.get(i);
            customerOrderTimeRptMapper.insertCustomerOrderTimeData(part);
            try {
                TimeUnit.MILLISECONDS.sleep(1000);
            } catch (InterruptedException e) {
                log.error("【CustomerOrderTimeRptService.saveCustomerOrderTimeToRptDB】客户工单时效写入中间表失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }
        }

    }

    /**
     * 删除中间表中指定日期的客户工单时效数据
     */
    private void deleteCustomerOrderTimeFromRptDB(Date date) {
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int systemId = RptCommonUtils.getSystemId();
            customerOrderTimeRptMapper.deleteCustomerOrderTimeData(systemId, beginDate.getTime(), endDate.getTime(),quarter);
        }
    }

    /**
     * 将工单系统中有的而中间表中没有的客户每日催单保存到中间表
     */
    public void saveMissedCustomerReminderToRptDB(Date date) {
        if (date != null) {
            List<CustomerOrderTimeRptEntity> list = getCustomerOrderTimeData(date);
            List<CustomerOrderTimeRptEntity> entityList = new ArrayList<>();
            if (!list.isEmpty()) {
                int systemId = RptCommonUtils.getSystemId();
                Map<Long, Long> completedOrderIdMap = getCompletedOrderIdMap(systemId, date);
                Long primaryKeyId;
                for (CustomerOrderTimeRptEntity entity : list) {
                    primaryKeyId = completedOrderIdMap.get(entity.getOrderId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        entityList.add(entity);

                    }
                }
                List<List<CustomerOrderTimeRptEntity>> parts = Lists.partition(entityList, 100);
                List<CustomerOrderTimeRptEntity> part;
                for (int i = 0, size = parts.size(); i < size; i++) {
                    part = parts.get(i);
                    customerOrderTimeRptMapper.insertCustomerOrderTimeData(part);
                    try {
                        TimeUnit.MILLISECONDS.sleep(1000);
                    } catch (InterruptedException e) {
                        log.error("【CustomerOrderTimeRptService.saveCustomerOrderTimeToRptDB】客户工单时效补漏写入中间表失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
                    }
                }

            }
        }
    }

    private Map<Long, Long> getCompletedOrderIdMap(Integer systemId, Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<LongTwoTuple> tuples = customerOrderTimeRptMapper.getCompletedOrderIds(quarter, systemId, beginDate.getTime(), endDate.getTime());
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }
    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveCustomerOrderTimeToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedCustomerReminderToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteCustomerOrderTimeFromRptDB(beginDate);
                            saveCustomerOrderTimeToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteCustomerOrderTimeFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
}
