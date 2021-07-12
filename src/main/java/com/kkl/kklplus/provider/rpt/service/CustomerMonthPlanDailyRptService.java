package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RptBaseMonthOrderEntity;
import com.kkl.kklplus.entity.rpt.RptCustomerMonthOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.CustomerMonthPlanDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
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
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerMonthPlanDailyRptService extends RptBaseService {
    @Autowired
    CustomerMonthPlanDailyRptMapper customerMonthPlanDailyRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    public List<RptCustomerMonthOrderEntity> getCustomerMonthPlanDailyList(RPTCustomerOrderPlanDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 1));
        int endYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 12));
        int systemId = RptCommonUtils.getSystemId();
        List<RptCustomerMonthOrderEntity> customerOrderPlanDailyList = new ArrayList<>();
        List<RptCustomerMonthOrderEntity> customerOrderPlanList = customerMonthPlanDailyRptMapper.getCustomerMonthPlanDailyList(startYearMonth, endYearMonth,search.getCustomerId(),search.getProductCategoryIds(),search.getSalesId(),systemId,search.getSubFlag());
        Map<Long, List<RptCustomerMonthOrderEntity>> orderDailyMap = customerOrderPlanList.stream().collect(Collectors.groupingBy(RptCustomerMonthOrderEntity::getCustomerId));

        double total;
        long customerId;
        Set<Long> customerIds = orderDailyMap.keySet();
        RPTCustomer customer;
        Set<Long> userIds = Sets.newHashSet();
        String[] fieldsArray = new String[]{"id", "name","code","salesId","paymentType","contractDate"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        userIds.addAll(salesIds);
        RPTDict paymentTypeDict;
        String userName;
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
        RptCustomerMonthOrderEntity rptEntity;
        try {
            for (List<RptCustomerMonthOrderEntity> entity : orderDailyMap.values()) {
                rptEntity = new RptCustomerMonthOrderEntity();
                total = 0;
                customerId = entity.get(0).getCustomerId();
                rptEntity.setCustomerId(customerId);
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);
                customer = customerMap.get(customerId);
                if(customer != null){
                    rptEntity.setCustomer(customer);
                    paymentTypeDict = paymentTypeMap.get(rptEntity.getCustomer().getPaymentType().getValue());
                    if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                        rptEntity.getCustomer().getPaymentType().setLabel(paymentTypeDict.getLabel());
                    }
                    userName = userNameMap.get(rptEntity.getCustomer().getSales().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        rptEntity.getCustomer().getSales().setName(userName);
                    }
                }else {
                    customer = new RPTCustomer();
                    rptEntity.setCustomer(customer);
                }
                rptEntity.setTotal(total);
                customerOrderPlanDailyList.add(rptEntity);
            }
        } catch (Exception e) {
            log.error("【CustomerMonthPlanDailyRptService.getCustomerMonthPlanDailyList】客户每日下单写入每日数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        customerOrderPlanDailyList = customerOrderPlanDailyList.stream().sorted(Comparator.comparing(i -> i.getCustomer().getName())).collect(Collectors.toList());
        RptCustomerMonthOrderEntity sumCOPM = new RptCustomerMonthOrderEntity();
        sumCOPM.setCustomer(new RPTCustomer());
        sumCOPM.setRowNumber(RptBaseMonthOrderEntity.RPT_ROW_NUMBER_SUMROW);
        sumCOPM.getCustomer().getSales().setName("总计(单)");

        RptCustomerMonthOrderEntity perCOPM = new RptCustomerMonthOrderEntity();
        perCOPM.setCustomer(new RPTCustomer());
        perCOPM.setRowNumber(RptBaseMonthOrderEntity.RPT_ROW_NUMBER_PERROW);
        perCOPM.getCustomer().getSales().setName("每月下单对比(%)");

        computeSumAndPerForCountOrAmount(customerOrderPlanDailyList, sumCOPM, perCOPM);
        customerOrderPlanDailyList.add(sumCOPM);
        customerOrderPlanDailyList.add(perCOPM);

        return customerOrderPlanDailyList;

    }

    /**
     * 写入每月的订单
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
    private double writeDailyOrders(RptCustomerMonthOrderEntity rptEntity, double total, Class rptEntityClass, List<RptCustomerMonthOrderEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int monthIndex;
        String dateStr;
        int month;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double monthSum;
        for (RptCustomerMonthOrderEntity entity : entityList) {
            monthIndex = entity.getYearMonth();
            monthSum = entity.getOrderSum();
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
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition.getStartDate() != null) {
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getStartDate()));
            int startYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 1));
            int endYearMonth = StringUtils.toInteger(String.format("%04d%02d", selectedYear, 12));
            Integer rowCount = customerMonthPlanDailyRptMapper.hasReportData(startYearMonth, endYearMonth,searchCondition.getCustomerId(),searchCondition.getProductCategoryIds(),searchCondition.getSalesId(),systemId,searchCondition.getSubFlag());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 获取客户每月下单图表

     * @return
     */
    public Map<String, Object> turnToCustomerOrderMonthPlanChart(RPTCustomerOrderPlanDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, 1, 1);
        int months = 12;
        List<Integer> monthSums = new ArrayList<>();
        List<String> createDates = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();

        List<RptCustomerMonthOrderEntity> copList = getCustomerMonthPlanDailyList(search);

        RptBaseMonthOrderEntity entity = copList.get(copList.size() - 2);
        copList.remove(copList.size()-1);
        copList.remove(copList.size()-1);
        copList =  copList.stream().sorted(Comparator.comparing(RptCustomerMonthOrderEntity::getTotal).reversed()).collect(Collectors.toList());
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
                    log.error("CustomerMonthPlanDailyRptService.turnToCustomerOrderMonthPlanChart",ex);
                }
            }
        }
        map.put("list",copList);
        map.put("createDates", createDates);
        map.put("monthSums", monthSums);
        return map;
    }


    /**
     * 客户每月派单报表导出
     *
     * @return
     */

    public SXSSFWorkbook customerOrderMonthRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RptCustomerMonthOrderEntity> list = getCustomerMonthPlanDailyList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 17));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算类型");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编码");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 16));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每月下单(单)及对比(%)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 17, 17));
            ExportExcel.createCell(headFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int monthIndex = 1; monthIndex <= 12; monthIndex++) {
                ExportExcel.createCell(headSecondRow, monthIndex + 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, monthIndex + "月");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RptCustomerMonthOrderEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getCustomer().getPaymentType() ? "" : rowData.getCustomer().getPaymentType().getLabel());

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getCustomer().getContractDate() ? "" : DateUtils.formatDate(rowData.getCustomer().getContractDate(), "yyyy-MM-dd"));

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getCustomer().getCode() ? "" : rowData.getCustomer().getCode());

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getCustomer().getName() ? "" : rowData.getCustomer().getName());

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getCustomer().getSales().getName() ? "" : rowData.getCustomer().getSales().getName());

                    Class rowDataClass = rowData.getClass();
                    for (int monthIndex = 1; monthIndex <= 12; monthIndex++) {
                        String strGetMethodName = "getM" + monthIndex;

                        Method method = rowDataClass.getMethod(strGetMethodName);
                        Object objGetM = method.invoke(rowData);
                        Double m = StringUtils.toDouble(objGetM);
                        String strM = null;
                        if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                            strM = String.format("%.2f%s", m, '%');
                        } else {
                            strM = String.format("%.0f", m);
                        }
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strM);
                    }

                    Double totalCount = StringUtils.toDouble(rowData.getTotal());
                    String strTotalCount = null;
                    if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
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
