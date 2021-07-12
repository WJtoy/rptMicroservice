package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTKefuCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.mapper.KeFuOrderPlanDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.KeFuUtils;
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
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuOrderPlanDailyRptService extends RptBaseService {
    @Resource
    private KeFuOrderPlanDailyRptMapper keFuOrderPlanDailyRptMapper;

    @Autowired
    private KeFuUtils keFuUtils;
    /**
     * 客服日接单报表
     *
     * @return
     */
    public List<RPTKeFuOrderPlanDailyEntity> getKeFuOrderAcceptDayRptData(RPTKeFuOrderPlanDailySearch search) {

        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        List<RPTKeFuOrderPlanDailyEntity> orderDailyPlanList = new ArrayList<>();

        Set<Long> keFuIds = new HashSet<>();
        Map<Long, RPTUser> kAKeFuMap = MSUserUtils.getMapByUserType(2);
        if(search.getSubFlag() != -1) {
            keFuUtils.getKeFu(kAKeFuMap,search.getSubFlag(),keFuIds);
        }
        List<Long> keFuIdsList =  Lists.newArrayList(keFuIds);
        Map<Long, RPTKeFuOrderPlanDailyEntity> keFusMap = new HashMap<>();
        List<RPTKeFuOrderPlanDailyEntity> orderDailyList = keFuOrderPlanDailyRptMapper.getKeFuOrderDailyList(search);


//        List<RPTKeFuOrderPlanDailyEntity> pendingQtyList = keFuOrderPlanDailyRptMapper.getKeFuPendingOrderQty(new Date(), search.getKeFuId(), search.getProductCategoryIds());
//
//
//        List<RPTKeFuOrderPlanDailyEntity> noGradedQtyList = keFuOrderPlanDailyRptMapper.getKeFuNoGradedOrderQty(search);


        Map<Long, List<RPTKeFuOrderPlanDailyEntity>> orderDailyMap = orderDailyList.stream().collect(Collectors.groupingBy(RPTKeFuOrderPlanDailyEntity::getKeFuId));

        double total;
        long keFuId;
        Set<Long> keFuIdSet = Sets.newHashSet();
        RPTKeFuOrderPlanDailyEntity rptEntity;
        try {
            for (List<RPTKeFuOrderPlanDailyEntity> entity : orderDailyMap.values()) {
                rptEntity = new RPTKeFuOrderPlanDailyEntity();
                total = 0;
                keFuId = entity.get(0).getKeFuId();
                rptEntity.setKeFuId(keFuId);
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);
                rptEntity.setTotal(total);
                orderDailyPlanList.add(rptEntity);
            }
        } catch (Exception e) {
            log.error("【KeFuOrderPlanDailyRptService.getKeFuOrderAcceptDayRptData】客服每日接单写入每日数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        Map<Long, RPTKeFuOrderPlanDailyEntity> orderDailyPlanMap = Maps.newHashMap();
        for (RPTKeFuOrderPlanDailyEntity item : orderDailyPlanList) {
            orderDailyPlanMap.put(item.getKeFuId(), item);
            keFuIdSet.add(item.getKeFuId());
        }
//        Map<Long, RPTKeFuOrderPlanDailyEntity> pendingQtyMap = Maps.newHashMap();
//        for (RPTKeFuOrderPlanDailyEntity item : pendingQtyList) {
//            pendingQtyMap.put(item.getKeFuId(), item);
//            keFuIdSet.add(item.getKeFuId());
//        }
//        Map<Long, RPTKeFuOrderPlanDailyEntity> noGradedQtyMap = Maps.newHashMap();
//        for (RPTKeFuOrderPlanDailyEntity item : noGradedQtyList) {
//            noGradedQtyMap.put(item.getKeFuId(), item);
//            keFuIdSet.add(item.getKeFuId());
//        }

        List<RPTKeFuOrderPlanDailyEntity> list = new ArrayList<>();
        RPTUser user;
        for (Long id : keFuIdSet) {
            RPTKeFuOrderPlanDailyEntity entity = orderDailyPlanMap.get(id);
            user = kAKeFuMap.get(id);
            if (entity == null) {
                entity = new RPTKeFuOrderPlanDailyEntity();
                entity.setKeFuId(id);
            } else {
                entity.setKeFuName(user == null ? "" : user.getName());
            }
//            RPTKeFuOrderPlanDailyEntity pendingQtyItem = pendingQtyMap.get(id);
//            if (pendingQtyItem != null) {
//                entity.setPendingQty(pendingQtyItem.getPendingQty());
//                entity.setKeFuName(user == null ? "" : user.getName());
//            }
//            RPTKeFuOrderPlanDailyEntity noGradedQtyItem = noGradedQtyMap.get(id);
//            if (noGradedQtyItem != null) {
//                entity.setNoGradedQty(noGradedQtyItem.getNoGradedQty());
//                entity.setKeFuName(user == null ? "" : user.getName());
//            }
            list.add(entity);
        }

        List<RPTKeFuOrderPlanDailyEntity> kAList = new ArrayList<>();
        if(search.getSubFlag() != -1){
            for(RPTKeFuOrderPlanDailyEntity item : list){
                keFusMap.put(item.getKeFuId(), item);
            }
            for(Long id : keFuIdsList){
               if(keFusMap.get(id) != null ){
                   kAList.add(keFusMap.get(id));
               }
            }
        }else{
            kAList.addAll(list);
        }
        kAList = kAList.stream().sorted(Comparator.comparing(RPTKeFuOrderPlanDailyEntity::getKeFuName)).collect(Collectors.toList());

        RPTKeFuOrderPlanDailyEntity sumKFOAD = new RPTKeFuOrderPlanDailyEntity();
        sumKFOAD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_SUMROW);
        sumKFOAD.setKeFuName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(kAList, 0, 0, sumKFOAD, null);
        kAList.add(sumKFOAD);

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
    private double writeDailyOrders(RPTKeFuOrderPlanDailyEntity rptEntity, double total, Class rptEntityClass, List<RPTKeFuOrderPlanDailyEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int dayIndex;
        String dateStr;
        int day;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double daySum;
        for (RPTKeFuOrderPlanDailyEntity entity : entityList) {
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
        RPTKeFuOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuOrderPlanDailySearch.class);
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
            Integer keFuOrderPlanQty = keFuOrderPlanDailyRptMapper.hasReportData(searchCondition.getSystemId(),searchCondition.getStartDate(),searchCondition.getEndDate(),keFuIdList,searchCondition.getProductCategoryIds(),searchCondition.getQuarter());
            Integer keFuPendingOrderQty = keFuOrderPlanDailyRptMapper.keFuPendingOrderQty(new Date(), keFuIdList, searchCondition.getProductCategoryIds());
            Integer keFuNoGradedOrderQty = keFuOrderPlanDailyRptMapper.keFuNoGradedOrderQty(keFuIdList, searchCondition.getProductCategoryIds());
            Integer rowCount = keFuOrderPlanQty + keFuPendingOrderQty + keFuNoGradedOrderQty;
            result = rowCount > 0;
        }
        return result;
    }


    /**
     * 客服每日接单报表导出
     *
     * @return
     */
    public SXSSFWorkbook keFuOrderPlanDailyRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTKeFuOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuOrderPlanDailySearch.class);
            List<RPTKeFuOrderPlanDailyEntity> list = getKeFuOrderAcceptDayRptData(searchCondition);

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
//            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
//            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "处理中");
//            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
//            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "停滞");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, days + 1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日接单(单)");

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
                int totalProcessingQty = 0;
                int totalPendingQty = 0;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuOrderPlanDailyEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getKeFuName() ? "" : rowData.getKeFuName());
//                    if (rowData.getRowNumber() != RPTBaseDailyEntity.RPT_ROW_NUMBER_SUMROW) {
//                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getProcessingQty());
//                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPendingQty());
//                    } else {
//                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalProcessingQty);
//                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPendingQty);
//                    }
//                    totalProcessingQty = totalProcessingQty + rowData.getProcessingQty();
//                    totalPendingQty = totalPendingQty + rowData.getPendingQty();

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
            log.error("【KeFuOrderPlanDailyRptService.keFuOrderAcceptDayRptExport】客服每日接单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
