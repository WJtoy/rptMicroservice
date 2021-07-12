package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderCancelledDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderCancelledDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.mapper.KeFuOrderCancelledDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
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
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuOrderCancelledDailyRptService extends RptBaseService{
    @Resource
    private KeFuOrderCancelledDailyRptMapper keFuOrderCancelledDailyRptMapper;

    @Autowired
    private KeFuUtils keFuUtils;
    /**
     * 客服日接单报表
     *
     * @return
     */
    public List<RPTKeFuOrderCancelledDailyEntity> getKeFuOrderCancelledDailyRptData(RPTKeFuOrderCancelledDailySearch search) {
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        List<RPTKeFuOrderCancelledDailyEntity> list = new ArrayList<>();

        Set<Long> keFuIds = new HashSet<>();
        Map<Long, RPTUser> kAKeFuMap =  MSUserUtils.getMapByUserType(2);
        if(search.getSubFlag() != -1) {
            keFuUtils.getKeFu(kAKeFuMap,search.getSubFlag(),keFuIds);
        }
        List<Long> keFuIdList =  Lists.newArrayList(keFuIds);
        Map<Long, RPTKeFuOrderCancelledDailyEntity> keFusMap = new HashMap<>();
        List<RPTKeFuOrderCancelledDailyEntity> entityList = keFuOrderCancelledDailyRptMapper.getKeFuCancelledOrderList(search);
        Map<Long, List<RPTKeFuOrderCancelledDailyEntity>> orderDailyMap = entityList.stream().collect(Collectors.groupingBy(RPTKeFuOrderCancelledDailyEntity::getKeFuId));
        double total;
        long keFuId;
        RPTUser user;
        RPTKeFuOrderCancelledDailyEntity rptEntity;
        try {
            for (List<RPTKeFuOrderCancelledDailyEntity> entity : orderDailyMap.values()) {
                rptEntity = new RPTKeFuOrderCancelledDailyEntity();
                total = 0;
                keFuId = entity.get(0).getKeFuId();
                rptEntity.setKeFuId(keFuId);
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);;
                if (kAKeFuMap != null && kAKeFuMap.size() > 0) {
                    user = kAKeFuMap.get(keFuId);
                    if (user != null) {
                        rptEntity.setKeFuName(user.getName());
                    }
                }
                rptEntity.setTotal(total);
                list.add(rptEntity);
            }
        } catch (Exception e) {
            log.error("【KeFuOrderCancelledDailyRptService.getKeFuOrderCancelledDailyRptData】客服每日退单写入每日数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        List<RPTKeFuOrderCancelledDailyEntity> kAList = new ArrayList<>();
        if(search.getSubFlag() != -1){
            for(RPTKeFuOrderCancelledDailyEntity item : list){
                keFusMap.put(item.getKeFuId(), item);
            }
            for(Long id : keFuIdList){
                if(keFusMap.get(id) != null ){
                    kAList.add(keFusMap.get(id));
                }
            }
        }else{
            kAList.addAll(list);
        }
        kAList = kAList.stream().sorted(Comparator.comparing(RPTKeFuOrderCancelledDailyEntity::getKeFuName)).collect(Collectors.toList());


        long startDt = search.getStartDate();

        Date startDate = new Date(startDt);
        Date lastMonthStartDate = DateUtils.addMonth(startDate, -1);
        String quarter = QuarterUtils.getSeasonQuarter(lastMonthStartDate);
        Date lastMonthEndDate = DateUtils.getLastDayOfMonth(lastMonthStartDate);
        search.setStartDate(lastMonthStartDate.getTime());
        search.setEndDate(lastMonthEndDate.getTime());
        search.setQuarter(quarter);
        int lastMonthDays = DateUtils.getDaysOfMonth(lastMonthStartDate);

        Long lastMonthOrderCount = keFuOrderCancelledDailyRptMapper.getKeFuOrderCancelledMonth(systemId,search.getStartDate(),search.getEndDate(),keFuIdList,search.getProductCategoryIds(),search.getQuarter());

        lastMonthOrderCount = (lastMonthOrderCount == null ? 0 : lastMonthOrderCount);
        double lastAvgCount = lastMonthOrderCount / lastMonthDays;

        RPTKeFuOrderCancelledDailyEntity sumKFORD = new RPTKeFuOrderCancelledDailyEntity();
        sumKFORD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_SUMROW);
        sumKFORD.setKeFuName("总计(单)");

        RPTKeFuOrderCancelledDailyEntity perKFORD = new RPTKeFuOrderCancelledDailyEntity();
        perKFORD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW);
        perKFORD.setKeFuName("对比(%)");

        RPTBaseDailyEntity.computeSumAndPerForCount(kAList, lastAvgCount, lastMonthOrderCount, sumKFORD, perKFORD);
        kAList.add(sumKFORD);
        kAList.add(perKFORD);
        //重新赋值导出时需要的当前查询月份
        search.setStartDate(startDt);

        return kAList;
    }

    /**
     * 写入每日的订单
     *
     * @param rptEntity
     * @param total
     * @param rptEntityClass
     * @param entityList
     * @return
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private double writeDailyOrders(RPTKeFuOrderCancelledDailyEntity rptEntity, double total, Class rptEntityClass, List<RPTKeFuOrderCancelledDailyEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int dayIndex;
        String dateStr;
        int day;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double daySum;
        for (RPTKeFuOrderCancelledDailyEntity entity : entityList) {
            dayIndex = entity.getDayIndex();
            daySum = entity.getOrderSum();
            if (dayIndex != 0) {
                dateStr = String.valueOf(dayIndex);
                day = StringUtils.toInteger(dateStr);
                strSetMethodName = "setD" + day;
                sumDailyReportSetMethod = rptEntityClass.getMethod(strSetMethodName, Double.class);
                total += daySum;
                sumDailyReportSetMethod.invoke(rptEntity, daySum);
            }

        }
        return total;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = false;
        RPTKeFuOrderCancelledDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuOrderCancelledDailySearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getStartDate() != null && searchCondition.getEndDate() != null) {
            Set<Long> keFuIds = new HashSet<>();
            Map<Long, RPTUser> kAKeFuMap;
            if(searchCondition.getSubFlag() != -1) {
                result = true;
                kAKeFuMap = MSUserUtils.getMapByUserType(2);
                keFuUtils.getKeFu(kAKeFuMap,searchCondition.getSubFlag(),keFuIds);
            }
            List<Long> keFuIdList =  Lists.newArrayList();
            List<Long> keFuSubTypeList =  Lists.newArrayList(keFuIds);
            if(null != searchCondition.getKeFuId() && result){
                if(keFuSubTypeList.size() > 0){
                    for(Long id : keFuSubTypeList){
                        if(searchCondition.getKeFuId().equals(id)){
                            keFuIdList.add(id);
                        }
                    }
                    if(keFuIdList.size() == 0){
                        return false;
                    }
                }else {
                    return false;
                }
            }else if(null != searchCondition.getKeFuId()){
                keFuIdList.add(searchCondition.getKeFuId());
            }else if(result){
                if(keFuSubTypeList.size() > 0){
                    keFuIdList.addAll(keFuSubTypeList);
                }else {
                    return false;
                }
            }
            Integer rowCount = keFuOrderCancelledDailyRptMapper.hasReportData(searchCondition.getSystemId(),searchCondition.getStartDate(),searchCondition.getEndDate(),keFuIdList,searchCondition.getProductCategoryIds(),searchCondition.getQuarter());
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook keFuOrderCancelledDailyRptExport(String searchConditionJson, String reportTitle) {


        SXSSFWorkbook xBook = null;
        try {
            RPTKeFuOrderCancelledDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuOrderCancelledDailySearch.class);
            List<RPTKeFuOrderCancelledDailyEntity> list = getKeFuOrderCancelledDailyRptData(searchCondition);

            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getStartDate()));

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 2));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, days + 1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日退单(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 2, days + 2));
            ExportExcel.createCell(headFirstRow, days + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuOrderCancelledDailyEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    if(dataRowIndex < rowsCount -2){
                        dataCell.setCellValue(dataRowIndex + 1);
                    }
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getKeFuName() ? "" : rowData.getKeFuName());

                    Class rowDataClass = rowData.getClass();
                    for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                        String strGetMethodName = "getD" + dayIndex;

                        Method method = rowDataClass.getMethod(strGetMethodName);
                        Object objGetD = method.invoke(rowData);
                        Double d = StringUtils.toDouble(objGetD);
                        String strD = null;
                        if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                            strD = String.format("%.2f%s", d, '%');
                        } else {
                            strD = String.format("%.0f", d);
                        }
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                    }

                    Double totalCount = StringUtils.toDouble(rowData.getTotal());
                    String strTotalCount = null;
                    if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                        strTotalCount = String.format("%.2f%s", totalCount, '%');

                    } else {
                        strTotalCount = String.format("%.0f", totalCount);
                    }
                     ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                }
            }


        } catch (Exception e) {
            log.error("【KeFuOrderCancelledDailyRptService.keFuOrderCancelledDailyRptExport】客服每日退单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
