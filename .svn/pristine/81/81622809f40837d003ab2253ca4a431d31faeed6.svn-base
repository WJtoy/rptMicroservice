package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.primitives.Ints;
import com.kkl.kklplus.entity.rpt.*;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.provider.rpt.entity.ChargeDailyEntity;
import com.kkl.kklplus.provider.rpt.mapper.ChargeDailyRptMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
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
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;
import java.util.Map;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ChargeDailyRptService extends RptBaseService {

    @Resource
    private ChargeDailyRptMapper chargeDailyRptMapper;

    public List<RPTChargeDailyEntity> getChargeDailyByList(RPTServicePointWriteOffSearch search) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
          Date min = new Date(search.getBeginWriteOffCreateDate());
          Date max = new Date(search.getEndWriteOffCreateDate());
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginWriteOffCreateDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getBeginWriteOffCreateDate()));
        int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
        int systemId = RptCommonUtils.getSystemId();
        //每日完成工单
        List<RPTAreaCompletedDailyEntity> keFuCompleteMonthList = chargeDailyRptMapper.getCompletedDaily(startYearMonth, systemId);
        int total = 0;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        Integer daySum;
        RPTChargeDailyEntity completedDaily = new RPTChargeDailyEntity();
        Class rptEntityClass = completedDaily.getClass();
        for(RPTAreaCompletedDailyEntity item : keFuCompleteMonthList){
            total = total + item.getCountSum();
            daySum = item.getCountSum();
            strSetMethodName = "setD" + item.getDayIndex();
            sumDailyReportSetMethod = rptEntityClass.getMethod(strSetMethodName, int.class);
            sumDailyReportSetMethod.invoke(completedDaily, daySum);

        }
        completedDaily.setRowSum(total);
        completedDaily.setChargeWay("完成工单");
        List<ChargeDailyEntity> orderDailyList 	= chargeDailyRptMapper.getManualChargeDaily(min , max);
        //手动对账
        RPTChargeDailyEntity chargeBaseEntity = new RPTChargeDailyEntity();
        chargeBaseEntity.setChargeWay("手动对账");
        int[] ints = new int[31];
        List<Integer> integers = Ints.asList(ints);
        int manual = 0;
        for (int i = 0; i < orderDailyList.size(); i++) {
            manual += orderDailyList.get(i).getCountNum().intValue();
        }
        for (int i = 0; i < orderDailyList.size(); i++) {
            for (int j = 0; j < integers.size(); j++) {
                if (orderDailyList.get(i).getDayIndex()==(j+1)){
                    integers.set(j,orderDailyList.get(i).getCountNum().intValue());
                }
            }

        }
        chargeBaseEntity.setRowSum(manual);
        chargeBaseEntity.setList(integers);
        //自动对账
        List<ChargeDailyEntity> autoChargeDaily = chargeDailyRptMapper.getAutoChargeDaily(min, max);
        RPTChargeDailyEntity chargeDailyEntity = new RPTChargeDailyEntity();
        chargeDailyEntity.setChargeWay("自动对账");
        int[] ints1 = new int[31];
        List<Integer> integers1 = Ints.asList(ints1);
        int auto = 0;
        for (int i = 0; i < autoChargeDaily.size(); i++) {
            auto += autoChargeDaily.get(i).getCountNum().intValue();
        }
        for (int i = 0; i < autoChargeDaily.size() ; i++) {
            for (int j = 0; j < integers1.size(); j++) {
                if (autoChargeDaily.get(i).getDayIndex()==(j+1)){
                    integers1.set(j,autoChargeDaily.get(i).getCountNum().intValue());
                }
            }
        }
        chargeDailyEntity.setRowSum(auto);
        chargeDailyEntity.setList(integers1);
        //合计(单)
        RPTChargeDailyEntity chargeDailyEntity1 = new RPTChargeDailyEntity();
        chargeDailyEntity1.setChargeWay("对账合计(单)");
        int[] ints2 = new int[31];
        List<Integer> integers2 = Ints.asList(ints2);
        for (int i = 0; i < integers2.size() ; i++) {
            integers2.set(i,integers.get(i)+integers1.get(i));
        }
        chargeDailyEntity1.setList(integers2);
        chargeDailyEntity1.setRowSum(auto+manual);
        List<RPTChargeDailyEntity> list = Lists.newArrayList();
        list.add(completedDaily);
        list.add(chargeBaseEntity);
        list.add(chargeDailyEntity);
        list.add(chargeDailyEntity1);
        return list;

    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
        if (searchCondition.getBeginWriteOffCreateDate() != null) {
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getBeginWriteOffCreateDate()));
            Integer selectedMonth = DateUtils.getMonth(new Date(searchCondition.getBeginWriteOffCreateDate()));
            int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
            int systemId = RptCommonUtils.getSystemId();
            Integer rowCount = chargeDailyRptMapper.getCompletedSum(startYearMonth,systemId);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook chargeDailyRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
            int selectedYear = DateUtils.getYear(new Date(searchCondition.getBeginWriteOffCreateDate()));
            int selectedMonth = DateUtils.getMonth(new Date(searchCondition.getBeginWriteOffCreateDate()));
            int days = DateUtils.getDaysOfMonth(DateUtils.parseDate(selectedYear + "-" + selectedMonth + "-01"));
            List<RPTChargeDailyEntity> list = getChargeDailyByList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);

            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 1));
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "统计名称");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, days));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 1, days + 1));
            ExportExcel.createCell(headFirstRow, days + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTChargeDailyEntity chargeDailyEntity = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == chargeDailyEntity.getChargeWay() ? "" : chargeDailyEntity.getChargeWay());

                    Class rowDataClass = chargeDailyEntity.getClass();
                    for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                        String strD = null;
                        if (dataRowIndex == 0) {
                            String strGetMethodName = "getD" + dayIndex;
                            Method method = rowDataClass.getMethod(strGetMethodName);
                            Object objGetD = method.invoke(chargeDailyEntity);
                            Double d = StringUtils.toDouble(objGetD);

                            strD = String.format("%.0f", d);
                        } else {
                            strD = String.valueOf(chargeDailyEntity.getList().get(dayIndex - 1));

                        }
                         ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                    }

                    Double totalCount = StringUtils.toDouble(chargeDailyEntity.getRowSum());
                    String strTotalCount = null;
                    if (chargeDailyEntity.getRowSum() == 100000) {
                        strTotalCount = String.format("%.2f%s", totalCount, '%');

                    } else {
                        strTotalCount = String.format("%.0f", totalCount);
                    }
                       ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                }
            }

        } catch (Exception e) {
            log.error("【ChargeDailyRptService.chargeDailyRptExport】每日对账统计报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}
