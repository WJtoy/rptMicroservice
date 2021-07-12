package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.*;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.mapper.KeFuCompleteMonthRptMapper;
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
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuCompletedMonthRptService extends  RptBaseService {

    @Autowired
    private KeFuCompleteMonthRptMapper keFuCompleteMonthRptMapper;

    @Autowired
    private KeFuUtils keFuUtils;

    public List<RPTKeFuCompletedMonthEntity> getKeFuCompletedMonthList(RPTGradedOrderSearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginDate()));
        int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 1));
        int endYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 12));
        int systemId = RptCommonUtils.getSystemId();

        Set<Long> keFuIds = new HashSet<>();
        Map<Long, RPTUser> kAKeFuMap  = MSUserUtils.getMapByUserType(2);
        if(search.getSubFlag() != -1) {
            keFuUtils.getKeFu(kAKeFuMap,search.getSubFlag(),keFuIds);
        }

        List<RPTKeFuCompletedMonthEntity> kefuCompletedMonthOrderList = new ArrayList<>();
        Map<Long,List<RPTKeFuCompletedMonthEntity>> keFusMap = new HashMap<>();
        List<RPTKeFuCompletedMonthEntity> keFuCompleteMonthList = keFuCompleteMonthRptMapper.getKeFuCompleteMonthList(startYearMonth, endYearMonth, search.getKefuId(), search.getProductCategoryIds(), systemId);
        Map<Long, List<RPTKeFuCompletedMonthEntity>> keFuMap = keFuCompleteMonthList.stream().collect(Collectors.groupingBy(RPTKeFuCompletedMonthEntity::getKefuId));
        if(search.getSubFlag() != -1){
            for(Long id : keFuIds){
                if(keFuMap.containsKey(id)){
                    keFusMap.put(id,keFuMap.get(id));
                }
            }
        }else{
            keFusMap = keFuMap;
        }
        RPTUser user;
        try {
            for (List<RPTKeFuCompletedMonthEntity> completedDailyEntityList : keFusMap.values()) {
                RPTKeFuCompletedMonthEntity completedDailyEntity = new RPTKeFuCompletedMonthEntity();

                if (kAKeFuMap != null && kAKeFuMap.size() > 0) {
                    user = kAKeFuMap.get(completedDailyEntityList.get(0).getKefuId());
                    if (user != null) {
                        completedDailyEntity.setKefuName(user.getName());
                    }
                }
                Double total = 0.0;
                Class rptEntityClass = completedDailyEntity.getClass();
                total = writeDailyOrders(completedDailyEntity, total, rptEntityClass, completedDailyEntityList);
                completedDailyEntity.setTotal(total);
                kefuCompletedMonthOrderList.add(completedDailyEntity);
            }
        } catch (Exception e) {
            log.error("【KeFuCompletedMonthRptService.getKeFuCompletedMonthList】客服完工单写入每月数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        RPTKeFuCompletedMonthEntity sumCOPM = new RPTKeFuCompletedMonthEntity();
        sumCOPM.setRowNumber(RptBaseMonthOrderEntity.RPT_ROW_NUMBER_SUMROW);
        sumCOPM.setKefuName("总计(单)");

        RPTKeFuCompletedMonthEntity perCOPM = new RPTKeFuCompletedMonthEntity();
        perCOPM.setRowNumber(RptBaseMonthOrderEntity.RPT_ROW_NUMBER_PERROW);
        perCOPM.setKefuName("每月完工单对比(%)");

        computeSumAndPerForCountOrAmount(kefuCompletedMonthOrderList, sumCOPM, perCOPM);
        kefuCompletedMonthOrderList.add(sumCOPM);
        kefuCompletedMonthOrderList.add(perCOPM);

        return kefuCompletedMonthOrderList;


    }

    private double writeDailyOrders(RPTKeFuCompletedMonthEntity rptEntity, double total, Class rptEntityClass, List<RPTKeFuCompletedMonthEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int monthIndex;
        String dateStr;
        int month;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double monthSum;
        for (RPTKeFuCompletedMonthEntity entity : entityList) {
            monthIndex = entity.getYearMonth();
            monthSum = entity.getCountSum();
            if (monthIndex != 0) {
                dateStr = String.valueOf(monthIndex % 100);
                month = StringUtils.toInteger(dateStr);
                strSetMethodName = "setM" + month;
                sumDailyReportSetMethod = rptEntityClass.getMethod(strSetMethodName, Double.class);
                total += monthSum;
                sumDailyReportSetMethod.invoke(rptEntity, monthSum);
            }

        }
        return total;
    }


    public static void computeSumAndPerForCountOrAmount(List baseMothlyRpts,
                                                        RptBaseMonthOrderEntity sumMothlyRpt, RptBaseMonthOrderEntity perMothlyRpt) {

        ////计算每月的总单数
        Class sumClass = sumMothlyRpt.getClass();
        for(Object object: baseMothlyRpts){
            RptBaseMonthOrderEntity item = (RptBaseMonthOrderEntity) object;
            Class itemClass = item.getClass();

            for(int i=1;i<13;i++) {
                String strGetMethodName ="getM"+i;
                String strSetMethodName ="setM"+i;
                try {
                    Method itemGetMethod   = itemClass.getMethod(strGetMethodName);
                    Method sumGetMethod    = sumClass.getMethod(strGetMethodName);

                    double mSum = StringUtils.toDouble(sumGetMethod.invoke(sumMothlyRpt))+ StringUtils.toDouble(itemGetMethod.invoke(item));

                    Method sumSetMethod = sumClass.getMethod(strSetMethodName, Double.class);
                    sumSetMethod.invoke(sumMothlyRpt,mSum);
                }
                catch (Exception ex) {
                    ex.printStackTrace();
                }
            }

            double mSumTotal 	= sumMothlyRpt.getTotal() + item.getTotal();
            sumMothlyRpt.setTotal(mSumTotal);
        }

        //计算每月订单数的对比值
        Class perClass = perMothlyRpt.getClass();
        for(int i=2;i<13;i++) {
            String strGetMethodName     ="getM"+(i-1);
            String strGetMethodNamePlus ="getM"+i;
            String strSetMethodNamePlus ="setM"+i;
            try {
                Method sumGetMethod     = sumClass.getMethod(strGetMethodName);
                Method sumGetMethodPlus = sumClass.getMethod(strGetMethodNamePlus);

                double sumM     = StringUtils.toDouble(sumGetMethod.invoke(sumMothlyRpt));
                double sumMPlus = StringUtils.toDouble(sumGetMethodPlus.invoke(sumMothlyRpt));

                double perM = -100;
                if (sumM != 0) {
                    perM =(sumMPlus/sumM-1) * 100;
                }

                Method perSetMethodPlus = perClass.getMethod(strSetMethodNamePlus, Double.class);
                perSetMethodPlus.invoke(perMothlyRpt, perM);
            }
            catch (Exception ex) {
                ex.printStackTrace();
            }
        }

    }

    /**
     * 获取客户每月完工图表

     * @return
     */
    public Map<String, Object> turnToKeFuCompletedMonthPlanChart(RPTGradedOrderSearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginDate()));
        Date queryDate = DateUtils.getDate(selectedYear, 1, 1);
        int months = 12;
        List<Integer> monthSums = new ArrayList<>();
        List<String> createDates = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();

        List<RPTKeFuCompletedMonthEntity> copList = getKeFuCompletedMonthList(search);

        RptBaseMonthOrderEntity entity = copList.get(copList.size() - 2);
        copList.remove(copList.size()-1);
        copList.remove(copList.size()-1);
        copList =  copList.stream().sorted(Comparator.comparing(RPTKeFuCompletedMonthEntity::getTotal).reversed()).collect(Collectors.toList());
        int size = copList.size();
        copList = copList.subList(0,size>10?10:size);
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        if (entity != null) {
            Class itemClass = entity.getClass();
            for (int i = 0; i < months; i++) {
                String Date = DateUtils.formatDate(DateUtils.addMonth(queryDate, i), "yyyy-MM");
                createDates.add(Date);
                String strGetMethodName = "getM" + (i + 1);
                try {
                    Method itemGetMethod = itemClass.getMethod(strGetMethodName);
                    Object itemGetM = itemGetMethod.invoke(entity);
                    Double dSum = StringUtils.toDouble(itemGetM);
                    int monthSum = (int) Math.floor(dSum);

                    monthSums.add(monthSum);

                } catch (Exception ex) {
                    log.error("KeFuCompletedMonthRptService.turnToKeFuCompletedMonthPlanChart",ex);
                }
            }
        }
        map.put("list",copList);
        map.put("createDates", createDates);
        map.put("monthSums", monthSums);
        return map;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = false;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition.getBeginDate() != null) {

            Set<Long> keFuIds = new HashSet<>();
            Map<Long, RPTUser> kAKeFuMap;
            if(searchCondition.getSubFlag() != -1) {
                result = true;
                kAKeFuMap = MSUserUtils.getMapByUserType(2);
                keFuUtils.getKeFu(kAKeFuMap,searchCondition.getSubFlag(),keFuIds);
            }
            List<Long> keFuIdsList =  Lists.newArrayList();
            List<Long> keFuSubTypeList =  Lists.newArrayList(keFuIds);
            if(searchCondition.getKefuId() != 0 && result){
                if(keFuSubTypeList.size() > 0){
                    for(Long id : keFuSubTypeList){
                        if(searchCondition.getKefuId().equals(id)){
                            keFuIdsList.add(id);
                        }
                    }
                    if(keFuIdsList.size() == 0){
                        return false;
                    }
                }else {
                    return false;
                }
            }else if(searchCondition.getKefuId() != 0){
                keFuIdsList.add(searchCondition.getKefuId());
            }else if(result){
                if(keFuSubTypeList.size() > 0){
                    keFuIdsList.addAll(keFuSubTypeList);
                }else {
                    return false;
                }
            }
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getBeginDate()));
            int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 1));
            int endYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 12));
            Integer rowCount = keFuCompleteMonthRptMapper.hasReportData(startYearMonth, endYearMonth, keFuIdsList, searchCondition.getProductCategoryIds(), systemId);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook keFuCompletedMonthRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            List<RPTKeFuCompletedMonthEntity> list = getKeFuCompletedMonthList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 12));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每月完工单(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 13, 13));
            ExportExcel.createCell(headFirstRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int monthIndex = 1; monthIndex <= 12; monthIndex++) {
                ExportExcel.createCell(headSecondRow, monthIndex, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, monthIndex + "月");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuCompletedMonthEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getKefuName() ? "" : rowData.getKefuName().toString());

                    Class rowDataClass = rowData.getClass();
                    for (int monthIndex = 1; monthIndex <= 12; monthIndex++) {
                        String strGetMethodName = "getM" + monthIndex;

                        Method method = rowDataClass.getMethod(strGetMethodName);
                        Object objGetM = method.invoke(rowData);
                        Double m = StringUtils.toDouble(objGetM);
                        String strM = null;
                        if (rowData.getRowNumber() == RptBaseMonthOrderEntity.RPT_ROW_NUMBER_PERROW) {
                            strM = String.format("%.2f%s", m, '%');
                        } else {
                            strM = String.format("%.0f", m);
                        }
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strM);
                    }

                    Double totalCount = StringUtils.toDouble(rowData.getTotal());
                    String strTotalCount = null;
                    if (rowData.getRowNumber() == RptBaseMonthOrderEntity.RPT_ROW_NUMBER_PERROW) {
                        strTotalCount = "";
                    } else {
                        strTotalCount = String.format("%.0f", totalCount);
                    }
                       ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                }
            }
        } catch (Exception e) {
            log.error("【CustomerMonthPlanDailyRptService.customerOrderMonthRptExport】客户每月下单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
