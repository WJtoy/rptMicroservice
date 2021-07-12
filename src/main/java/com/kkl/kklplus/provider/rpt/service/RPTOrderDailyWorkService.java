package com.kkl.kklplus.provider.rpt.service;

import com.kkl.kklplus.entity.b2bcenter.md.B2BDataSourceEnum;
import com.kkl.kklplus.entity.rpt.RPTOrderDailyWorkEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.provider.rpt.mapper.RptCreatedOrderMapper;
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
public class RPTOrderDailyWorkService extends  RptBaseService {

    @Autowired
    RptCreatedOrderMapper rptCreatedOrderMapper;

    public List<RPTOrderDailyWorkEntity> getCreatedOrderList(RPTCompletedOrderDetailsSearch search) {
        search.setSystemId(RptCommonUtils.getSystemId());
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getEndDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Date startDate = DateUtils.getStartDayOfMonth(queryDate);
        List<RPTOrderDailyWorkEntity> entityList = rptCreatedOrderMapper.getOrderDailyWorkList(search);
        List<RPTOrderDailyWorkEntity> list = new ArrayList<>();
        int days = DateUtils.getDaysOfMonth(queryDate);
        Date date = new Date();
        int year = DateUtils.getYear(date);
        int month = DateUtils.getMonth(date);
        int day = Integer.valueOf(DateUtils.getDay());

        if(year == selectedYear && month < selectedMonth){
            return list;
        }


        int manualOrderSum = 0;
        int tmOrderSum = 0;
        int jdOrderSum = 0;
        int restOrderSum = 0;
        int pddOrderSum = 0;
        int daySum = 0;

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        if (year == selectedYear && month == selectedMonth) {
            days = day;
        }

        Map<Integer, List<RPTOrderDailyWorkEntity>> groupBy = entityList.stream().collect(Collectors.groupingBy(RPTOrderDailyWorkEntity::getDayIndex));

        String key;
        String yearStr;
        String monthStr;
        String dayStr;
        for (int i = 0; i < days; i++) {
            Integer keyInt = new Integer(DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd"));
            RPTOrderDailyWorkEntity rptEntity = new RPTOrderDailyWorkEntity();
            int sum = 0;
            key = String.valueOf(keyInt);
            yearStr = key.substring(0, 4);
            monthStr = key.substring(4, 6);
            dayStr = key.substring(6, 8);
            if (groupBy.get(i+1) == null) {
                rptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + dayStr);
            } else {
                rptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + dayStr);
                for (RPTOrderDailyWorkEntity entity : groupBy.get(i+1)) {
                    Integer dataSource = entity.getDataSource();
                    if (dataSource != null) {
                        if (dataSource == B2BDataSourceEnum.KKL.id) {
                            rptEntity.setManualOrder(entity.getDataSourceSum());
                        } else if (dataSource == B2BDataSourceEnum.TMALL.id) {
                            rptEntity.setTmOrder(entity.getDataSourceSum());
                        } else if (dataSource == B2BDataSourceEnum.JD.id) {
                            rptEntity.setJdOrder(entity.getDataSourceSum());
                        } else if(dataSource == B2BDataSourceEnum.PDD.id){
                            rptEntity.setPddOrder(entity.getDataSourceSum());
                        } else {
                            if (entity.getDataSourceSum() != null) {
                                rptEntity.setRestOrder(rptEntity.getRestOrder() + entity.getDataSourceSum());
                            }
                        }
                        sum += entity.getDataSourceSum();
                    }
                }
                rptEntity.setDaySum(sum);
                Integer total = rptEntity.getDaySum();
                countOrderSourceRate(numberFormat, rptEntity, total);
                manualOrderSum += rptEntity.getManualOrder();
                tmOrderSum += rptEntity.getTmOrder();
                jdOrderSum += rptEntity.getJdOrder();
                restOrderSum += rptEntity.getRestOrder();
                pddOrderSum += rptEntity.getPddOrder();
                daySum += rptEntity.getDaySum();
            }
            list.add(rptEntity);
        }
        list = list.stream().sorted(Comparator.comparing(RPTOrderDailyWorkEntity::getOrderCreateDate)).collect(Collectors.toList());
        //合计
        RPTOrderDailyWorkEntity orderDailyWorkRptEntity = new RPTOrderDailyWorkEntity();
        orderDailyWorkRptEntity.setManualOrder(manualOrderSum);
        orderDailyWorkRptEntity.setTmOrder(tmOrderSum);
        orderDailyWorkRptEntity.setJdOrder(jdOrderSum);
        orderDailyWorkRptEntity.setRestOrder(restOrderSum);
        orderDailyWorkRptEntity.setPddOrder(pddOrderSum);
        orderDailyWorkRptEntity.setDaySum(daySum);
        Integer total = orderDailyWorkRptEntity.getDaySum();
        countOrderSourceRate(numberFormat, orderDailyWorkRptEntity, total);
        list.add(orderDailyWorkRptEntity);

        return list;

    }

    /**
     * 计算每日工单来源的比率
     *
     * @param numberFormat
     * @param rptEntity
     * @param total
     */
    private void countOrderSourceRate(NumberFormat numberFormat, RPTOrderDailyWorkEntity rptEntity, Integer total) {
        if (total != null && total != 0) {
            int manualOrder = rptEntity.getManualOrder();
            if (manualOrder > 0) {
                String manualOrderRete = numberFormat.format((float) manualOrder / total * 100);
                rptEntity.setManualOrderRate(manualOrderRete);
            }
            int tmOrder = rptEntity.getTmOrder();
            if (tmOrder > 0) {
                String tmOrderRete = numberFormat.format((float) tmOrder / total * 100);
                rptEntity.setTmOrderOrderRate(tmOrderRete);
            }
            int jdOrder = rptEntity.getJdOrder();
            if (jdOrder > 0) {
                String jdOrderRete = numberFormat.format((float) jdOrder / total * 100);
                rptEntity.setJdOrderRate(jdOrderRete);
            }
            int pddOrder = rptEntity.getPddOrder();
            if(pddOrder >0){
                String pddOrderRete = numberFormat.format((float) pddOrder / total * 100);
                rptEntity.setPddOrderRate(pddOrderRete);
            }
            int restOrder = rptEntity.getRestOrder();
            if (restOrder > 0) {
                String restOrderRete = numberFormat.format((float) restOrder / total * 100);
                rptEntity.setRestOrderRate(restOrderRete);
            }
        }
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginDate() != null && searchCondition.getEndDate() != null) {
            int OrderDailyWorkSum = rptCreatedOrderMapper.getOrderDailyWorkSum(searchCondition);
            result = OrderDailyWorkSum > 0;
        }
        return result;
    }


    public SXSSFWorkbook OrderDailyWorkExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
            List<RPTOrderDailyWorkEntity> list = getCreatedOrderList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手动工单");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手动工单比率");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "天猫工单");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "天猫工单比率");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "拼多多工单");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "拼多多工单比率");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "京东工单");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "京东工单比率");
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他工单");
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他工单比率");
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "总数量");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTOrderDailyWorkEntity OrderDailyWorkRptEntity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    if (dataRowIndex == list.size() - 1) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                    } else {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getOrderCreateDate());
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getManualOrder());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getManualOrderRate() + "%");
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getTmOrder());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getTmOrderOrderRate() + "%");
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getPddOrder());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getPddOrderRate() + "%");
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getJdOrder());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getJdOrderRate() + "%");
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getRestOrder());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getRestOrderRate() + "%");
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, OrderDailyWorkRptEntity.getDaySum());
                }
            }
        } catch (Exception e) {
            log.error("【RPTOrderDailyWorkService.OrderDailyWorkExport】工单来源统计报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

}
