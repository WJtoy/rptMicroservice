package com.kkl.kklplus.provider.rpt.service;

import com.kkl.kklplus.entity.rpt.RPTDispatchOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.mapper.DispatchListRptMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
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

import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class DispatchListInformationRptService  extends RptBaseService {

    @Autowired
    private DispatchListRptMapper dispatchListRptMapper;

    public List<RPTDispatchOrderEntity> getDispatchListInformation(RPTCompletedOrderDetailsSearch search){
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getEndDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Date startDate = DateUtils.getStartDayOfMonth(queryDate);
        Date date = new Date();
        int year = DateUtils.getYear(date);
        int month = DateUtils.getMonth(date);
        List<RPTDispatchOrderEntity> list = new ArrayList<>();

        if(year == selectedYear && month < selectedMonth){
             return  list;
        }


        List<RPTDispatchOrderEntity> planInformationSums = dispatchListRptMapper.getPlanInformationSum(search);

        List<RPTDispatchOrderEntity> theTotalOrders = dispatchListRptMapper.getTheTotalOrder(search);

        List<RPTDispatchOrderEntity> cancelOrderSums = dispatchListRptMapper.getCancelOrderSum(search);

        if (theTotalOrders == null || theTotalOrders.size() <= 0) {
            return list;
        }
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        int days = DateUtils.getDaysOfMonth(queryDate);
        int day  = Integer.valueOf(DateUtils.getDay());

        if(year == selectedYear && month == selectedMonth){
            days = day-1;
        }

        int notPlanSum = 0;
        int autoSum = 0;
        int appSum = 0;
        int keFuSum = 0;
        int crushSum = 0;
        int cancelSum = 0;
        int planSum = 0;
        int theTotalOrderSum = 0;
        Map<Integer, List<RPTDispatchOrderEntity>> groupBy = planInformationSums.stream().collect(Collectors.groupingBy(RPTDispatchOrderEntity::getDayIndex));

        Map<Integer, List<RPTDispatchOrderEntity>> theTotalOrderMap = theTotalOrders.stream().collect(Collectors.groupingBy(RPTDispatchOrderEntity::getDayIndex));

        Map<Integer, List<RPTDispatchOrderEntity>> cancelOrderSumMap = cancelOrderSums.stream().collect(Collectors.groupingBy(RPTDispatchOrderEntity::getDayIndex));

        String key;
        String yearStr;
        String monthStr;
        String dayStr;
        for(int i = 0;i < days; i++){
            Integer keyInt = new Integer(DateUtils.formatDate(DateUtils.addDays(startDate, i),"yyyyMMdd"));
            RPTDispatchOrderEntity rptEntity = new RPTDispatchOrderEntity();
            key = String.valueOf(keyInt);
            yearStr = key.substring(0, 4);
            monthStr = key.substring(4, 6);
            dayStr = key.substring(6, 8);
            if(groupBy.get(keyInt) == null){
                rptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + dayStr);
            }else {
                rptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + dayStr);
                for (RPTDispatchOrderEntity entity : groupBy.get(keyInt)) {
                    Integer planType = entity.getPlanType();
                    if(planType != null){
                        if(planType == 2){
                            rptEntity.setApp(entity.getPlanTypeSum());
                        }else if(planType == 1){
                            rptEntity.setAuto(entity.getPlanTypeSum());
                        }else if(planType == 3){
                            rptEntity.setKefu(entity.getPlanTypeSum());
                        }else if(planType == 4){
                            rptEntity.setCrush(entity.getPlanTypeSum());
                        }else if(planType == 0){
                            rptEntity.setNotPlan(entity.getPlanTypeSum());
                        }
                    }

                }
                if(theTotalOrderMap.get(keyInt) != null){
                    for(RPTDispatchOrderEntity entity : theTotalOrderMap.get(keyInt)){
                        rptEntity.setTheTotalOrder(entity.getTheTotalOrder());
                    }
                }
                if(cancelOrderSumMap.get(keyInt) != null){
                    for(RPTDispatchOrderEntity entity :cancelOrderSumMap.get(keyInt)){
                        rptEntity.setCancel(entity.getCancel());
                    }
                }
                rptEntity.getSum();
                Integer total = rptEntity.getTheTotalOrder();
                countPlanInformationRptRate(numberFormat, rptEntity, total);
                notPlanSum += rptEntity.getNotPlan();
                autoSum += rptEntity.getAuto();
                appSum += rptEntity.getApp();
                keFuSum += rptEntity.getKefu();
                crushSum += rptEntity.getCrush();
                cancelSum += rptEntity.getCancel();
                planSum += rptEntity.getSum() ;
                theTotalOrderSum += rptEntity.getTheTotalOrder();
            }
            list.add(rptEntity);

        }
        list = list.stream().sorted(Comparator.comparing(RPTDispatchOrderEntity::getOrderCreateDate)).collect(Collectors.toList());
        //合计
        RPTDispatchOrderEntity planInformationRptEntity = new RPTDispatchOrderEntity();
        planInformationRptEntity.setNotPlan(notPlanSum);
        planInformationRptEntity.setAuto(autoSum);
        planInformationRptEntity.setApp(appSum);
        planInformationRptEntity.setKefu(keFuSum);
        planInformationRptEntity.setCrush(crushSum);
        planInformationRptEntity.setCancel(cancelSum);
        planInformationRptEntity.setSum(planSum);
        planInformationRptEntity.setTheTotalOrder(theTotalOrderSum);

        Integer total = planInformationRptEntity.getTheTotalOrder();
        countPlanInformationRptRate(numberFormat, planInformationRptEntity, total);
        list.add(planInformationRptEntity);

        return list;

    }

    /**
     * 计算每日派单报表中的比率
     * @param numberFormat
     * @param rptEntity
     * @param total
     */
    private void countPlanInformationRptRate(NumberFormat numberFormat, RPTDispatchOrderEntity rptEntity, Integer total) {
        if (total != null && total != 0) {
            int auto = rptEntity.getAuto();
            if(auto > 0){
                String autoRate = numberFormat.format((float) auto / total * 100);
                rptEntity.setAutoRate(autoRate);
            }
            int app = rptEntity.getApp();
            if(app > 0){
                String appRate = numberFormat.format((float) app / total * 100);
                rptEntity.setAppRate(appRate);
            }
            int keFu = rptEntity.getKefu();
            if(keFu > 0){
                String kefuRate = numberFormat.format((float) keFu / total * 100);
                rptEntity.setKefuRate(kefuRate);
            }
            int crush = rptEntity.getCrush();
            if(crush > 0){
                String crushRate = numberFormat.format((float) crush / total * 100);
                rptEntity.setCrushRate(crushRate);
            }
            int notPlan = rptEntity.getNotPlan();
            if(notPlan > 0){
                String notPlanRate = numberFormat.format((float) notPlan / total * 100);
                rptEntity.setNotPlanRate(notPlanRate);
            }
            int cancelSum = rptEntity.getCancel();
            if(cancelSum > 0){
                String cancelSumRate = numberFormat.format((float) cancelSum / total * 100);
                rptEntity.setCancelRate(cancelSumRate);
            }
            int planTypeSum = rptEntity.getPlanTypeSum();
            if(planTypeSum > 0){
                String planTypeSumRate = numberFormat.format((float) planTypeSum / total * 100);
                rptEntity.setPlanTypeSumRate(planTypeSumRate);
            }
        }
    }

    public Map<String, Object> getDispatchListInformationChart(RPTCompletedOrderDetailsSearch search) {
        Map<String, Object> map = new HashMap<>();
        List<RPTDispatchOrderEntity> entityList = getDispatchListInformation(search);
        if (entityList == null || entityList.size() <= 0) {
            return map;
        }


        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getEndDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int days = DateUtils.getDaysOfMonth(queryDate);
        int day  = Integer.valueOf(DateUtils.getDay());

        RPTDispatchOrderEntity planInformationRptEntity = entityList.get(entityList.size() - 1);

        Map<String, String> planTypeSumMap = new HashMap<>();
        Map<String, String> autoMap = new HashMap<>();
        Map<String, String> appMap = new HashMap<>();
        Map<String, String> keFuMap = new HashMap<>();
        Map<String, String> crushMap = new HashMap<>();
        Map<String, String> cancelMap = new HashMap<>();
        Map<String, String> notPlanMap = new HashMap<>();

        List<Map<String, String>> mapList = new ArrayList<>();

        List<Map<String, String>> planList = new ArrayList<>();

        planTypeSumMap.put("value",planInformationRptEntity.getSum().toString());
        planTypeSumMap.put("name","已派单");
        autoMap.put("value",planInformationRptEntity.getAuto().toString());
        autoMap.put("name","自动派单");
        appMap.put("value",planInformationRptEntity.getApp().toString());
        appMap.put("name","APP派单");
        keFuMap.put("value",planInformationRptEntity.getKefu().toString());
        keFuMap.put("name","客服派单");
        crushMap.put("value",planInformationRptEntity.getCrush().toString());
        crushMap.put("name","突击单");
        notPlanMap.put("value",planInformationRptEntity.getNotPlan().toString());
        notPlanMap.put("name","未派单");
        cancelMap.put("value",planInformationRptEntity.getCancel().toString());
        cancelMap.put("name","取消单");


        List<String> createDates = new ArrayList<>();
        List<String> orderCreateDates = new ArrayList<>();

        List<String> strAuto = new ArrayList<>();
        List<String> strApp = new ArrayList<>();
        List<String> strKeFu = new ArrayList<>();
        List<String> strCrush = new ArrayList<>();
        List<String> strNotPlan = new ArrayList<>();
        List<String> strCancel = new ArrayList<>();


        List<String> strAutoRate = new ArrayList<>();
        List<String> strAppRate = new ArrayList<>();
        List<String> strKeFuRate = new ArrayList<>();
        List<String> strCrushRate = new ArrayList<>();
        List<String> strNotPlanRate = new ArrayList<>();
        List<String> strCancelRate = new ArrayList<>();
        int year;
        int month;
        for(RPTDispatchOrderEntity entity : entityList){
            if(entity.getOrderCreateDate()!=null) {

                orderCreateDates.add(entity.getOrderCreateDate().substring(8));
                createDates.add(entity.getOrderCreateDate().substring(8));

                strAuto.add(entity.getAuto().toString());
                strApp.add(entity.getApp().toString());
                strKeFu.add(entity.getKefu().toString());
                strCrush.add(entity.getCrush().toString());
                strNotPlan.add(entity.getNotPlan().toString());
                strCancel.add(entity.getCancel().toString());

                strAutoRate.add(entity.getAutoRate());
                strAppRate.add(entity.getAppRate());
                strKeFuRate.add(entity.getKefuRate());
                strCrushRate.add(entity.getCrushRate());
                strNotPlanRate.add(entity.getNotPlanRate());
                strCancelRate.add(entity.getCancelRate());
            }
        }
        year = DateUtils.getYear(new Date());
        month = DateUtils.getMonth(new Date());
        if(selectedYear == year && selectedMonth == month){
            for(int i = day ;i < days;i++){
                String orderCreateDate = DateUtils.formatDate(DateUtils.addDays(queryDate, i), "dd");
                orderCreateDates.add(orderCreateDate);
            }
        }

        planList.add(planTypeSumMap);
        planList.add(cancelMap);
        planList.add(notPlanMap);

        mapList.add(autoMap);
        mapList.add(appMap);
        mapList.add(keFuMap);
        mapList.add(crushMap);
        mapList.add(cancelMap);
        mapList.add(notPlanMap);


        map.put("mapList",mapList);
        map.put("planList",planList);

        map.put("createDates", createDates);
        map.put("orderCreateDates", orderCreateDates);
        map.put("strAuto",strAuto);
        map.put("strApp",strApp);
        map.put("strKeFu",strKeFu);
        map.put("strCrush",strCrush);
        map.put("strNotPlan",strNotPlan);
        map.put("strCancel",strCancel);

        map.put("strAutoRate",strAutoRate);
        map.put("strAppRate",strAppRate);
        map.put("strKeFuRate",strKeFuRate);
        map.put("strCrushRate",strCrushRate);
        map.put("strNotPlanRate",strNotPlanRate);
        map.put("strCancelRate",strCancelRate);


        return map;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
       searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginDate()!= null && searchCondition.getEndDate() != null) {
            Integer rowCount = dispatchListRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook DisPatchListInfoRptExport(String searchConditionJson, String reportTitle) {
          SXSSFWorkbook xBook = null;

        try {
            RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
            List<RPTDispatchOrderEntity> list = getDispatchListInformation(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单数量");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单数量");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "自动派单");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "自动派单比率");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "App抢单");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "App抢单比率");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服派单");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服派单比率");
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "突击单");
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "突击单比率");
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未派单");
            ExportExcel.createCell(headFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未派单比率");
            ExportExcel.createCell(headFirstRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消单");
            ExportExcel.createCell(headFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消单比率");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTDispatchOrderEntity entity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    if (dataRowIndex == list.size() - 1) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                    } else {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderCreateDate());
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTheTotalOrder());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getSum());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAuto());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAutoRate() + "%");
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getApp());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAppRate() + "%");
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getKefu());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getKefuRate() + "%");
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCrush());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCrushRate() + "%");
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getNotPlan());
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getNotPlanRate() + "%");
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCancel());
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCancelRate() + "%");

                }
            }
        } catch (Exception e) {
            log.error("【DispatchListInformationRptService.DisPatchListInfoRptExport】客户每月下单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
          return xBook;
    }

}
