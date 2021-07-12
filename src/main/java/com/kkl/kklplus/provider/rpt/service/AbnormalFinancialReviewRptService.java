package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.abnormal.AbnormalFinancialAuditStatistics;
import com.kkl.kklplus.entity.rpt.RPTAbnormalFinancialAuditEntity;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.entity.AbnormalFinancialReviewEntity;
import com.kkl.kklplus.provider.rpt.mapper.AbnormalFinancialReviewRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
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

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class AbnormalFinancialReviewRptService extends RptBaseService {

    @Autowired
    private AbnormalFinancialReviewRptMapper abnormalFinancialReviewRptMapper;

    public List<RPTAbnormalFinancialAuditEntity> getAbnormalFinancialAuditList(RPTCustomerOrderPlanDailySearch search){
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));

        List<RPTAbnormalFinancialAuditEntity> abnormalFinancialList = new ArrayList<>();

        List<RPTAbnormalFinancialAuditEntity> abnormalSumList = new ArrayList<>();

        List<RPTAbnormalFinancialAuditEntity> abnormalManualChargeList = new ArrayList<>();

        if(search.getId() == 0) {

            double completedTotal = 0;
            double autoChargeTotal = 0;


            List<RPTAbnormalFinancialAuditEntity>   completedDailyList = abnormalFinancialReviewRptMapper.getCompletedDaily(startYearMonth, systemId);

            List<RPTAbnormalFinancialAuditEntity>  autoChargeList = abnormalFinancialReviewRptMapper.getAutoMonthList(search);       //自动

            RPTAbnormalFinancialAuditEntity completedEntity = new RPTAbnormalFinancialAuditEntity();
            RPTAbnormalFinancialAuditEntity autoChargeEntity = new RPTAbnormalFinancialAuditEntity();

            completedEntity.setCreateName("完成工单");
            autoChargeEntity.setCreateName("自动对账");


            Class completedEntityClass = completedEntity.getClass();
            Class autoChargeEntityClass = autoChargeEntity.getClass();

            try {
                completedTotal = writeDailyOrders(completedEntity, completedTotal, completedEntityClass, completedDailyList);
                autoChargeTotal = writeDailyOrders(autoChargeEntity, autoChargeTotal, autoChargeEntityClass, autoChargeList);
                completedEntity.setTotal(completedTotal);
                autoChargeEntity.setTotal(autoChargeTotal);

                abnormalFinancialList.add(completedEntity);
                abnormalFinancialList.add(autoChargeEntity);

            }catch (Exception e){
                log.error("【AbnormalFinancialReviewRptService.getAbnormalFinancialAuditList】财务审单数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }



        }

        List<RPTAbnormalFinancialAuditEntity> manualChargeList = abnormalFinancialReviewRptMapper.getManualMonthList(search);       //手动

        List<RPTAbnormalFinancialAuditEntity> abnormalList = abnormalFinancialReviewRptMapper.getAbnormalList(search);

        Map<Integer, List<RPTAbnormalFinancialAuditEntity>> abnormalDailyMap = abnormalList.stream().collect(Collectors.groupingBy(RPTAbnormalFinancialAuditEntity::getSubType));



        Map<String, RPTDict> abnormalMap = MSDictUtils.getDictMap("fi_charge_audit_type");//切换为微服务

        double manualChargeTotal = 0;
        double abnormalChargeTotal;
        Integer subTypeId;
        RPTDict rptDictSubType;
        RPTAbnormalFinancialAuditEntity manualChargeEntity = new RPTAbnormalFinancialAuditEntity();
        RPTAbnormalFinancialAuditEntity abnormalEntity;

        Class manualChargeEntityClass = manualChargeEntity.getClass();

        if(search.getId() == 0 ){
            manualChargeEntity.setCreateName("手动对账（合计）");
        }else {
            manualChargeEntity.setCreateName("手动对账");
        }
        try {
            manualChargeTotal = writeDailyOrders(manualChargeEntity, manualChargeTotal, manualChargeEntityClass, manualChargeList);
            manualChargeEntity.setTotal(manualChargeTotal);
            abnormalFinancialList.add(manualChargeEntity);
            abnormalManualChargeList.add(manualChargeEntity);

            for(List<RPTAbnormalFinancialAuditEntity> abnormalFinancialAuditEntityList : abnormalDailyMap.values()){
                    abnormalEntity = new RPTAbnormalFinancialAuditEntity();
                    subTypeId = abnormalFinancialAuditEntityList.get(0).getSubType();
                    abnormalChargeTotal = 0;
                    Class abnormalEntityClass = abnormalEntity.getClass();
                    abnormalChargeTotal = writeDailyOrders(abnormalEntity, abnormalChargeTotal, abnormalEntityClass, abnormalFinancialAuditEntityList);
                    rptDictSubType = abnormalMap.get(String.valueOf(subTypeId));
                    if(rptDictSubType != null && StringUtils.isNotBlank(rptDictSubType.getLabel())){
                        abnormalEntity.setCreateName(rptDictSubType.getLabel());
                     }

                    abnormalEntity.setTotal(abnormalChargeTotal);

                    abnormalFinancialList.add(abnormalEntity);
                    abnormalSumList.add(abnormalEntity);
                    abnormalManualChargeList.add(abnormalEntity);

            }


        } catch (Exception e){
            log.error("【AbnormalFinancialReviewRptService.getAbnormalFinancialAuditList】财务审单数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        RPTAbnormalFinancialAuditEntity abnormalSum = new RPTAbnormalFinancialAuditEntity();
        abnormalSum.setCreateName("异常工单（合计）");
        RPTBaseDailyEntity.computeSumAndPerForCount(abnormalSumList, 0, 0, abnormalSum, null);
        abnormalFinancialList.add(abnormalSum);

        RPTAbnormalFinancialAuditEntity sumUp = new RPTAbnormalFinancialAuditEntity();
        sumUp.setCreateName("手动对账+异常单（合计）");
        RPTBaseDailyEntity.computeSumAndPerForCount(abnormalManualChargeList, 0, 0, sumUp, null);
        abnormalFinancialList.add(sumUp);

       return abnormalFinancialList;

    }


    public Map<String,List<RPTAbnormalFinancialAuditEntity>> getCheckerList(RPTCustomerOrderPlanDailySearch search){
        int systemId = RptCommonUtils.getSystemId();

        List<RPTAbnormalFinancialAuditEntity> getAbnormalFinancialAuditList = getAbnormalFinancialAuditList(search);

        Map<String,List<RPTAbnormalFinancialAuditEntity>> checkerMap =  new LinkedHashMap<>();

        Map<Long, String> checkersMap;

        Set<Long> abnormalIds = Sets.newHashSet();

        String name = "统计名称";

        if(search.getId() != 0){
            abnormalIds.add(search.getId());
            checkersMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(abnormalIds));
            String checkerName = checkersMap.get(search.getId());
            name = "审单员数据(" + checkerName + ")";
        }
        checkerMap.put(name,getAbnormalFinancialAuditList);


        if(search.getId() == 0) {

            List<RPTAbnormalFinancialAuditEntity> abnormalFinancialList;

            List<RPTAbnormalFinancialAuditEntity> abnormalSumList;

            List<RPTAbnormalFinancialAuditEntity> abnormalManualChargeList;


            List<RPTAbnormalFinancialAuditEntity> abnormalList = abnormalFinancialReviewRptMapper.getCheckerAbnormalList(search);

            for (RPTAbnormalFinancialAuditEntity abnormalItem : abnormalList) {
                abnormalIds.add(abnormalItem.getCreateId());
            }

            List<RPTAbnormalFinancialAuditEntity> manualChargeList = abnormalFinancialReviewRptMapper.getCheckerManualList(search.getStartDate(), search.getEndDate(), systemId, Lists.newArrayList(abnormalIds)); //审单员手动

            Map<Long, List<RPTAbnormalFinancialAuditEntity>> checkerManualMap = manualChargeList.stream().collect(Collectors.groupingBy(RPTAbnormalFinancialAuditEntity::getCreateId));

            Map<Long, List<RPTAbnormalFinancialAuditEntity>> checkerAbnormalMap = abnormalList.stream().collect(Collectors.groupingBy(RPTAbnormalFinancialAuditEntity::getCreateId));


            Map<String, RPTDict> abnormalMap = MSDictUtils.getDictMap("fi_charge_audit_type");
            checkersMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(abnormalIds));

            Integer subTypeId;
            RPTDict rptDictSubType;
            RPTAbnormalFinancialAuditEntity manualChargeEntity;
            RPTAbnormalFinancialAuditEntity abnormalEntity;
            RPTAbnormalFinancialAuditEntity abnormalSum;
            RPTAbnormalFinancialAuditEntity sumUp;
            double abnormalChargeTotal;
            try {
                for (Long abnormalId : abnormalIds) {
                    double manualChargeTotal = 0;
                    abnormalManualChargeList = new ArrayList<>();
                    abnormalFinancialList = new ArrayList<>();
                    abnormalSumList = new ArrayList<>();
                    manualChargeEntity = new RPTAbnormalFinancialAuditEntity();
                    manualChargeEntity.setCreateName("手动对账（合计）");
                    Class manualChargeEntityClass = manualChargeEntity.getClass();

                    List<RPTAbnormalFinancialAuditEntity> checkerManualList = checkerManualMap.get(abnormalId);
                    if (checkerManualList != null && checkerManualList.size() > 0) {
                        manualChargeTotal = writeDailyOrders(manualChargeEntity, manualChargeTotal, manualChargeEntityClass, checkerManualList);
                        manualChargeEntity.setTotal(manualChargeTotal);
                    }
                    abnormalFinancialList.add(manualChargeEntity);

                    List<RPTAbnormalFinancialAuditEntity> checkerAbnormalList = checkerAbnormalMap.get(abnormalId);

                    if (checkerAbnormalList != null && checkerAbnormalList.size() > 0) {
                        Map<Integer, List<RPTAbnormalFinancialAuditEntity>> abnormalDailyMap = checkerAbnormalList.stream().collect(Collectors.groupingBy(RPTAbnormalFinancialAuditEntity::getSubType));
                        for (List<RPTAbnormalFinancialAuditEntity> abnormalFinancialAuditEntityList : abnormalDailyMap.values()) {
                            abnormalEntity = new RPTAbnormalFinancialAuditEntity();
                            subTypeId = abnormalFinancialAuditEntityList.get(0).getSubType();
                            abnormalChargeTotal = 0;
                            Class abnormalEntityClass = abnormalEntity.getClass();
                            abnormalChargeTotal = writeDailyOrders(abnormalEntity, abnormalChargeTotal, abnormalEntityClass, abnormalFinancialAuditEntityList);
                            rptDictSubType = abnormalMap.get(String.valueOf(subTypeId));
                            if (rptDictSubType != null && StringUtils.isNotBlank(rptDictSubType.getLabel())) {
                                abnormalEntity.setCreateName(rptDictSubType.getLabel());
                            }

                            abnormalEntity.setTotal(abnormalChargeTotal);
                            abnormalFinancialList.add(abnormalEntity);
                            abnormalSumList.add(abnormalEntity);
                        }
                    }


                    abnormalSum = new RPTAbnormalFinancialAuditEntity();
                    abnormalSum.setCreateName("异常工单（合计）");
                    RPTBaseDailyEntity.computeSumAndPerForCount(abnormalSumList, 0, 0, abnormalSum, null);

                    abnormalManualChargeList.add(manualChargeEntity);
                    abnormalManualChargeList.add(abnormalSum);
                    sumUp = new RPTAbnormalFinancialAuditEntity();
                    sumUp.setCreateName("手动对账+异常单（合计）");
                    RPTBaseDailyEntity.computeSumAndPerForCount(abnormalManualChargeList, 0, 0, sumUp, null);

                    abnormalFinancialList.add(abnormalSum);
                    abnormalFinancialList.add(sumUp);


                    String checkerName = checkersMap.get(abnormalId);
                    checkerMap.put("审单员数据(" + checkerName + ")", abnormalFinancialList);

                }

            } catch (Exception e) {
                log.error("【AbnormalFinancialReviewRptService.getCheckerList】财务审单数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }

        }
         return checkerMap;


    }



    private double writeDailyOrders(RPTAbnormalFinancialAuditEntity rptEntity, double total, Class rptEntityClass, List<RPTAbnormalFinancialAuditEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int dayIndex;
        String dateStr;
        int day;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double daySum;
        for (RPTAbnormalFinancialAuditEntity entity : entityList) {
            dayIndex = entity.getDayIndex();
            daySum = entity.getCountSum();
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

    private List<AbnormalFinancialReviewEntity> getAbnormalFinancialList(Date startDate, Date endDate){

        int day = Integer.valueOf(DateUtils.getDay(startDate));

        int systemId = RptCommonUtils.getSystemId();

        List<AbnormalFinancialReviewEntity>  list = new ArrayList<>();

        List<AbnormalFinancialReviewEntity>  manualChargeList =  abnormalFinancialReviewRptMapper.getManualChargeList(startDate,endDate);  //手动

        List<AbnormalFinancialReviewEntity>  autoChargeList =  abnormalFinancialReviewRptMapper.getAutoChargeList(startDate,endDate);   //自动

        List<AbnormalFinancialAuditStatistics>  abnormalList  =   abnormalFinancialReviewRptMapper.financialAuditStat(startDate.getTime(),endDate.getTime());


        AbnormalFinancialReviewEntity entity;
        for(AbnormalFinancialReviewEntity  manualItem :manualChargeList){
            manualItem.setDayIndex(day);
            manualItem.setType(20);
            manualItem.setAuditTime(startDate.getTime());
            manualItem.setSystemId(systemId);
            list.add(manualItem);
        }

        for(AbnormalFinancialReviewEntity  autoItem :autoChargeList){
            autoItem.setDayIndex(day);
            autoItem.setType(10);
            autoItem.setAuditTime(startDate.getTime());
            autoItem.setSystemId(systemId);
            list.add(autoItem);
        }

        for(AbnormalFinancialAuditStatistics  abnormalItem :abnormalList){
            entity = new AbnormalFinancialReviewEntity();
            entity.setSubType(abnormalItem.getSubType());
            entity.setCreateId(abnormalItem.getCreateId());
            entity.setDayIndex(day);
            entity.setQty(abnormalItem.getQty());
            entity.setType(30);
            entity.setProductCategoryId(abnormalItem.getProductCategoryId());
            entity.setAuditTime(startDate.getTime());
            entity.setSystemId(systemId);
            list.add(entity);
        }

        return list;
    }


    public void saveAbnormalFinancialReviewRptDB(Date date) {
             saveAbnormalFinancialRptDB(date);
    }

    public void saveAbnormalFinancialRptDB(Date date){
        if(date != null){
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            List <AbnormalFinancialReviewEntity> list = getAbnormalFinancialList(startDate,endDate);
            for(AbnormalFinancialReviewEntity item : list){
                abnormalFinancialReviewRptMapper.insertAbnormalFinancialDaily(item);
            }
        }
    }

    private void deleteAbnormalFinancialRptDB(Date date){
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            abnormalFinancialReviewRptMapper.deleteAbnormalFinancialData(beginDate.getTime(), endDate.getTime(), systemId);
        }
    }


    public boolean rebuildMiddleAbnormalTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveAbnormalFinancialRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteAbnormalFinancialRptDB(beginDate);
                            saveAbnormalFinancialRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteAbnormalFinancialRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("AbnormalFinancialReviewRptService.rebuildMiddleAbnormalTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;

    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch search = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        Integer rowCount;
        if (search.getStartDate() != null && search.getEndDate() != null) {
            Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
            Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
            int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
            if(search.getId() == 0){
              rowCount = abnormalFinancialReviewRptMapper.hasReportData(startYearMonth,systemId);
            }else{
               rowCount = abnormalFinancialReviewRptMapper.hasManualReportData(search);
            }

            result = rowCount > 0;
        }
        return result;
    }



    public SXSSFWorkbook abnormalPlanDailyRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            Map<String,List<RPTAbnormalFinancialAuditEntity>> checkerMap = getCheckerList(searchCondition);
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getStartDate()));
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 2));

            //====================================================绘制表头============================================================

            for (Map.Entry<String, List<RPTAbnormalFinancialAuditEntity>> m : checkerMap.entrySet()) {
                String titleName  = m.getKey();
                List<RPTAbnormalFinancialAuditEntity> list = m.getValue();

                Row headFirstRow = xSheet.createRow(rowIndex++);
                headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
                xSheet.addMergedRegion(new CellRangeAddress(rowIndex-1, rowIndex, 0, 0));
                ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

                xSheet.addMergedRegion(new CellRangeAddress(rowIndex-1, rowIndex, 1, 1));
                ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, titleName);


                xSheet.addMergedRegion(new CellRangeAddress(rowIndex-1, rowIndex-1, 2, days + 1));
                ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日(单)");

                xSheet.addMergedRegion(new CellRangeAddress(rowIndex-1, rowIndex, days + 2, days + 2));
                ExportExcel.createCell(headFirstRow, days + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

                Row headSecondRow = xSheet.createRow(rowIndex++);
                headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                    ExportExcel.createCell(headSecondRow, dayIndex + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
                }

                int j = 1;
                Cell dataCell = null;
                if (list != null && list.size() > 0) {
                    int rowsCount = list.size();
                    for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                        RPTAbnormalFinancialAuditEntity rowData = list.get(dataRowIndex);

                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 0;
                        dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        dataCell.setCellValue(dataRowIndex + 1);

                        dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        dataCell.setCellValue(null == rowData.getCreateName() ? "" : rowData.getCreateName());

                        Class rowDataClass = rowData.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetMethodName = "getD" + dayIndex;

                            Method method = rowDataClass.getMethod(strGetMethodName);
                            Object objGetD = method.invoke(rowData);
                            Double d = StringUtils.toDouble(objGetD);
                            String strD;
                            if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                                strD = String.format("%.2f%s", d, '%');
                            } else {
                                strD = String.format("%.0f", d);
                            }
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                        }

                        Double totalCount = StringUtils.toDouble(rowData.getTotal());
                        String strTotalCount;
                        if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                            strTotalCount = String.format("%.2f%s", totalCount, '%');

                        } else {
                            strTotalCount = String.format("%.0f", totalCount);
                        }
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                    }
                }
                if(titleName.equals("统计名称")){
                    xSheet.createRow(rowIndex=rowIndex+2);

                }else{
                    xSheet.createRow(rowIndex++);
                }

            }

        } catch (Exception e) {
            log.error("【AbnormalFinancialReviewRptService.abnormalPlanDailyRptExport】客户每日下单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }



}
