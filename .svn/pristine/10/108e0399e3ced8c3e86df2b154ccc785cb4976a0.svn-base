package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTGradeQtyDailyEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.entity.CustomerPerformanceRptEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyRptEntity;
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;
import com.kkl.kklplus.provider.rpt.mapper.GradeQtyDailyRptMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
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
public class GradeQtyDailyRptService extends RptBaseService{


    @Autowired
    private GradeQtyDailyRptMapper gradeQtyDailyRptMapper;

     public List<RPTGradeQtyDailyEntity> getGradeQtyRptList(RPTCustomerOrderPlanDailySearch search) {
        List<RPTGradeQtyDailyEntity> list = new ArrayList<>();
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getEndDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int days = DateUtils.getDaysOfMonth(queryDate);
        int day = Integer.valueOf(DateUtils.getDay());
        Date date = new Date();
        int year = DateUtils.getYear(date);
        int month = DateUtils.getMonth(date);

         if(year == selectedYear && month < selectedMonth){
             return list;
         }

         if (year == selectedYear && month == selectedMonth) {
             days = day-1 ;
         }

         list = gradeQtyDailyRptMapper.getGradeQtyDailyList(search);
         Map<Integer, List<RPTGradeQtyDailyEntity>> groupBy = list.stream().collect(Collectors.groupingBy(RPTGradeQtyDailyEntity::getDayIndex));

        int manualGradeQtySum = 0;
        int appGradeQtySum = 0;
        int smsGradeQtySum = 0;
        int voiceGradeQtySum = 0;
        int totalQtySum = 0;
        List<RPTGradeQtyDailyEntity> lists = Lists.newArrayList();


        for (RPTGradeQtyDailyEntity gradeQtyEntity : list) {
            if (gradeQtyEntity.getTotalQty() != 0) {

                manualGradeQtySum += gradeQtyEntity.getManualGradeQty();
                appGradeQtySum += gradeQtyEntity.getAppGradeQty();
                smsGradeQtySum += gradeQtyEntity.getSmsGradeQty();
                voiceGradeQtySum += gradeQtyEntity.getVoiceGradeQty();
                totalQtySum += gradeQtyEntity.getTotalQty();
                getGradeQtyPercentage(gradeQtyEntity);
            }
            lists.add(gradeQtyEntity);
        }

         for (int i = 0; i < days; i++) {
              if(groupBy.get(i+1) == null){
                  RPTGradeQtyDailyEntity gradeQtyEntity = new RPTGradeQtyDailyEntity();
                  gradeQtyEntity.   setDayIndex(i+1);
                  lists.add(gradeQtyEntity);

              }
         }

         lists = lists.stream().sorted(Comparator.comparing(RPTGradeQtyDailyEntity::getDayIndex)).collect(Collectors.toList());

        RPTGradeQtyDailyEntity gradeQtyEntity = new RPTGradeQtyDailyEntity();
        gradeQtyEntity.setManualGradeQty(manualGradeQtySum);
        gradeQtyEntity.setAppGradeQty(appGradeQtySum);
        gradeQtyEntity.setSmsGradeQty(smsGradeQtySum);
        gradeQtyEntity.setVoiceGradeQty(voiceGradeQtySum);
        gradeQtyEntity.setTotalQty(totalQtySum);

        if (gradeQtyEntity.getTotalQty() != 0) {
            getGradeQtyPercentage(gradeQtyEntity);
            String totalQtyPercentage = 100 + "%";
            gradeQtyEntity.setTotalQtyPercentage(totalQtyPercentage);
            lists.add(gradeQtyEntity);
        }

          return  lists;
    }

    public RPTGradeQtyDailyEntity getGradeQtyPercentage(RPTGradeQtyDailyEntity gradeQtyEntity){

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        String manualGradeQtyPercentage;
        String appGradeQtyPercentage;
        String smsGradeQtyPercentage;
        String voiceGradeQtyPercentage;

        manualGradeQtyPercentage = numberFormat.format((float) gradeQtyEntity.getManualGradeQty() / gradeQtyEntity.getTotalQty() * 100);
        appGradeQtyPercentage = numberFormat.format((float) gradeQtyEntity.getAppGradeQty() / gradeQtyEntity.getTotalQty() * 100);
        smsGradeQtyPercentage = numberFormat.format((float) gradeQtyEntity.getSmsGradeQty() / gradeQtyEntity.getTotalQty() * 100);
        voiceGradeQtyPercentage = numberFormat.format((float) gradeQtyEntity.getVoiceGradeQty() / gradeQtyEntity.getTotalQty() * 100);

        gradeQtyEntity.setManualGradeQtyPercentage(manualGradeQtyPercentage);
        gradeQtyEntity.setAppGradeQtyPercentage(appGradeQtyPercentage);
        gradeQtyEntity.setSmsGradeQtyPercentage(smsGradeQtyPercentage);
        gradeQtyEntity.setVoiceGradeQtyPercentage(voiceGradeQtyPercentage);

        return gradeQtyEntity;

    }



    public List<GradeQtyRptEntity> getGradeQty(Date startDate,  Date endDate) {
        List<GradeQtyEntity> gradeQtyList = gradeQtyDailyRptMapper.getGradeQtyList(startDate, endDate);
        int day = Integer.valueOf(DateUtils.getDay(startDate));
        int systemId = RptCommonUtils.getSystemId();
        List<GradeQtyRptEntity> gradeList = new ArrayList<>();
        GradeQtyRptEntity gradeQtyRptEntity ;
        Map<Long, List<GradeQtyEntity>> gradeQtyMap = gradeQtyList.stream().collect(Collectors.groupingBy(GradeQtyEntity::getProductCategoryId));
        Set<Long> productCategoryIds = Sets.newHashSet();
        for (GradeQtyEntity entity : gradeQtyList) {
            productCategoryIds.add(entity.getProductCategoryId());
        }
        for (long productCategoryId : productCategoryIds) {
            int total = 0;
            int manualTotal = 0;
            int smsTotal = 0;
            int voiceTotal = 0;
            int appTotal = 0;
            gradeQtyRptEntity = new GradeQtyRptEntity();
            List<GradeQtyEntity> list = gradeQtyMap.get(productCategoryId);
            for (GradeQtyEntity qtyEntity : list) {
                if (qtyEntity.getGradeFlag() == 1) {
                    manualTotal = manualTotal + qtyEntity.getCount();
                }
                if (qtyEntity.getGradeFlag() == 2) {
                    smsTotal = smsTotal + qtyEntity.getCount();
                }
                if (qtyEntity.getGradeFlag() == 3) {
                    voiceTotal = voiceTotal + qtyEntity.getCount();
                }
                if (qtyEntity.getGradeFlag() == 4) {
                    appTotal = appTotal + qtyEntity.getCount();
                }
            }


            gradeQtyRptEntity.setSystemId(systemId);
            gradeQtyRptEntity.setManualGradeQty(manualTotal);
            gradeQtyRptEntity.setSmsGradeQty(smsTotal);
            gradeQtyRptEntity.setAppGradeQty(appTotal);
            gradeQtyRptEntity.setVoiceGradeQty(voiceTotal);
            gradeQtyRptEntity.setProductCategoryId(productCategoryId);
            total = manualTotal + smsTotal + appTotal + voiceTotal;
            gradeQtyRptEntity.setTotalQty(total);
            gradeQtyRptEntity.setGradeDate(startDate.getTime());
            gradeQtyRptEntity.setDayIndex(day);
            gradeList.add(gradeQtyRptEntity);
        }

        return  gradeList;
    }

    private Map<String, Long> getGradeIdMap(int systemId, Date startDate, Date endDate) {

        List<LongThreeTuple> tuples = gradeQtyDailyRptMapper.getGradeQtyDailyIds(systemId,startDate.getTime(), endDate.getTime());
        Map<String, Long> tuplesMap = Maps.newHashMap();
        if(tuples != null && !tuples.isEmpty()) {
            for(LongThreeTuple item : tuples ){
                String key = StringUtils.join(item.getBElement(), "%", item.getCElement());
                tuplesMap.put(key,item.getAElement());
            }
            return tuplesMap;
        } else {
            return tuplesMap;
        }
    }

    /**
     * 写入当前一天数据
     */
    public void writeYesterdayQty(Date date) {
        updateCustomerPerformanceToRptDB(date);

    }

    private void updateCustomerPerformanceToRptDB(Date date) {
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<GradeQtyRptEntity> list = getGradeQty(startDate, endDate);
        if (!list.isEmpty()) {
            int systemId = RptCommonUtils.getSystemId();
            Map<String, Long> idMap = getGradeIdMap(systemId, startDate,endDate);
            Long primaryKeyId;
            String key;
            for (GradeQtyRptEntity item : list) {
                key = StringUtils.join(item.getGradeDate(), "%", item.getProductCategoryId());
                primaryKeyId = idMap.get(key);
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    gradeQtyDailyRptMapper.updateGradeQty(item);
                } else {
                    gradeQtyDailyRptMapper.insertGradeQty(item);
                }
            }
        }
    }

    public void deleteGradeQtyRptDB(Date date){
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            gradeQtyDailyRptMapper.delete(beginDate.getTime(), endDate.getTime(), systemId);
        }
    }

    public void saveGradeQtyRptDB(Date date){
        updateCustomerPerformanceToRptDB(date);

    }

    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveGradeQtyRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            updateCustomerPerformanceToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteGradeQtyRptDB(beginDate);
                            saveGradeQtyRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteGradeQtyRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("GradeQtyDailyRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;

    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        int systemId = RptCommonUtils.getSystemId();
        searchCondition.setSystemId(systemId);
        if (searchCondition.getStartDate() != null) {
            Integer rowCount = gradeQtyDailyRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook gradeQtyDailyRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTGradeQtyDailyEntity> gradeQtyRpt = getGradeQtyRptList(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "统计项目");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手动客评");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 1, 2));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "APP完成客评");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 3, 4));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "短信客评");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 5, 6));
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "语音客评");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum(), 7, 8));
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "总客评单量");
            xSheet.addMergedRegion(new CellRangeAddress(headFirstRow.getRowNum(), headFirstRow.getRowNum() + 1, 9, 9));
            //表头第二行===============================
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headerSecondRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "天");
            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单量");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评占比");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单量");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评占比");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单量");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评占比");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单量");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评占比");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            if (gradeQtyRpt != null && gradeQtyRpt.size() > 0) {
                int rowsCount = gradeQtyRpt.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                        RPTGradeQtyDailyEntity gradeQtyEntity = gradeQtyRpt.get(dataRowIndex);
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                        if (dataRowIndex == gradeQtyRpt.size() - 1) {
                            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                        } else {
                            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getDayIndex());
                        }
                            ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getManualGradeQty());
                            ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getManualGradeQtyPercentage()+"%");
                            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getAppGradeQty());
                            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getAppGradeQtyPercentage()+"%");
                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getSmsGradeQty());
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getSmsGradeQtyPercentage()+"%");
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getVoiceGradeQty());
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getVoiceGradeQtyPercentage()+"%");
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, gradeQtyEntity.getTotalQty());


                    }
                }


        } catch (Exception e) {
            log.error("【GradeQtyDailyRptService.customerOrderMonthRptExport】客评统计报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}


