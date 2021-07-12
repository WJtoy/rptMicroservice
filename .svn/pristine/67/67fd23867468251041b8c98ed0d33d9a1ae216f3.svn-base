package com.kkl.kklplus.provider.rpt.customer.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.CustomerOrderPlanDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
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
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCustomerOrderPlanDailyRptService extends RptBaseService {
    @Resource
    private CustomerOrderPlanDailyRptMapper customerOrderPlanDailyRptMapper;
    @Autowired
    private MSCustomerService msCustomerService;

    public List<RPTCustomerOrderPlanDailyEntity> getCustomerOrderPlanDailyRptData(RPTCustomerOrderPlanDailySearch search) {
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        List<RPTCustomerOrderPlanDailyEntity> customerOrderPlanDailyList = new ArrayList<>();
        List<RPTCustomerOrderPlanDailyEntity> customerOrderPlanList = customerOrderPlanDailyRptMapper.getCustomerOrderPlanDailyList(search);
        Map<Long, List<RPTCustomerOrderPlanDailyEntity>> orderDailyMap = customerOrderPlanList.stream().collect(Collectors.groupingBy(RPTCustomerOrderPlanDailyEntity::getCustomerId));

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
        RPTCustomerOrderPlanDailyEntity rptEntity;
        try {
            for (List<RPTCustomerOrderPlanDailyEntity> entity : orderDailyMap.values()) {
                rptEntity = new RPTCustomerOrderPlanDailyEntity();
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
            log.error("【CustomerOrderPlanDailyRptService.getCustomerOrderPlanDailyRptData】客户每日下单写入每日数据错误, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        }

        long startDt = search.getStartDate();
        Date startDate = new Date(startDt);

        Date lastMonthStartDate = DateUtils.addMonth(startDate, -1);
        String quarter = QuarterUtils.getSeasonQuarter(lastMonthStartDate);
        Date lastMonthEndDate = DateUtils.getLastDayOfMonth(lastMonthStartDate);
        search.setQuarter(quarter);
        search.setStartDate(lastMonthStartDate.getTime());
        search.setEndDate(lastMonthEndDate.getTime());
        int lastMonthDays = DateUtils.getDaysOfMonth(lastMonthStartDate);

        Long lastMonthCount = customerOrderPlanDailyRptMapper.getCustomerOrderPlanMonth(search);
        lastMonthCount = (lastMonthCount == null ? 0 : lastMonthCount);
        double lastMonthAvgCount = (lastMonthCount * 1.0) / lastMonthDays;
        customerOrderPlanDailyList = customerOrderPlanDailyList.stream().sorted(Comparator.comparing(i -> i.getCustomer().getName())).collect(Collectors.toList());
        RPTCustomerOrderPlanDailyEntity sumCOPD = new RPTCustomerOrderPlanDailyEntity();
        sumCOPD.setCustomer(new RPTCustomer());
        sumCOPD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_SUMROW);
        sumCOPD.getCustomer().getSales().setName("总计(单)");

        RPTCustomerOrderPlanDailyEntity perCOPD = new RPTCustomerOrderPlanDailyEntity();
        perCOPD.setCustomer(new RPTCustomer());
        perCOPD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW);
        perCOPD.getCustomer().getSales().setName("每日下单环比(%)");

        RPTBaseDailyEntity.computeSumAndPerForCount(customerOrderPlanDailyList, lastMonthAvgCount, lastMonthCount, sumCOPD, perCOPD);

        RPTCustomerOrderPlanDailyEntity perSCOPD = new RPTCustomerOrderPlanDailyEntity();
        Date goLiveDate = RptCommonUtils.getGoLiveDate();
        Date lastYearStartDate = DateUtils.addMonth(startDate, -12);
        quarter = QuarterUtils.getSeasonQuarter(lastYearStartDate);
        Date lastYearEndDate = DateUtils.getLastDayOfMonth(lastYearStartDate);
        if (goLiveDate.getTime() < lastYearStartDate.getTime()) {
            search.setStartDate(lastYearStartDate.getTime());
            search.setEndDate(lastYearEndDate.getTime());
            search.setQuarter(quarter);
            int lastYearSomeMonthDays = DateUtils.getDaysOfMonth(lastYearStartDate);

            Long lastYearSomeMonthCount = customerOrderPlanDailyRptMapper.getCustomerOrderPlanMonth(search);
            lastYearSomeMonthCount = (lastYearSomeMonthCount == null ? 0 : lastYearSomeMonthCount);
            double lastYearSomeMonthAvgCount = (lastYearSomeMonthCount * 1.0) / lastYearSomeMonthDays;
            computeSumAndPerForCount(customerOrderPlanDailyList, lastYearSomeMonthAvgCount, lastYearSomeMonthCount, sumCOPD, perSCOPD);
        }
        perSCOPD.setCustomer(new RPTCustomer());
        perSCOPD.setRowNumber(RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW);
        perSCOPD.getCustomer().getSales().setName("每日下单同比(%)");
        customerOrderPlanDailyList.add(sumCOPD);
        customerOrderPlanDailyList.add(perCOPD);
        customerOrderPlanDailyList.add(perSCOPD);

        //重新赋值导出时需要的当前查询月份
        search.setStartDate(startDt);
        return customerOrderPlanDailyList;
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
    private double writeDailyOrders(RPTCustomerOrderPlanDailyEntity rptEntity, double total, Class rptEntityClass, List<RPTCustomerOrderPlanDailyEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int dayIndex;
        String dateStr;
        int day;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double daySum;
        for (RPTCustomerOrderPlanDailyEntity entity : entityList) {
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
     * 计算每日每月订单数的比率
     *
     * @param baseDailyReports
     * @param lastMonthAvgCount
     * @param lastMonthTotalCount
     * @param sumDailyReport
     * @param perDailyReport
     */
    public static void computeSumAndPerForCount(List baseDailyReports, double lastMonthAvgCount, double lastMonthTotalCount,
                                                RPTBaseDailyEntity sumDailyReport, RPTBaseDailyEntity perDailyReport) {
        //计算每月订单数的对比值
        if (sumDailyReport != null && perDailyReport != null) {
            Class sumDailyReportClass = sumDailyReport.getClass();
            Class perDailyReportClass = perDailyReport.getClass();
            for (int i = 1; i < 32; i++) {
                String strGetMethodName = "getD" + i;
                String strSetMethodName = "setD" + i;
                try {
                    Method sumDailyReportGetMethod = sumDailyReportClass.getMethod(strGetMethodName);
                    Object sumDailyReportGetD = sumDailyReportGetMethod.invoke(sumDailyReport);

                    double sumD = StringUtils.toDouble(sumDailyReportGetD);
                    double perD = -100;
                    if (lastMonthAvgCount != 0) {
                        perD = (sumD - lastMonthAvgCount) / lastMonthAvgCount * 100;
                    }

                    Method perDailyReportSetMethod = perDailyReportClass.getMethod(strSetMethodName, Double.class);
                    perDailyReportSetMethod.invoke(perDailyReport, perD);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }

            double perTotal = -100;
            if (lastMonthTotalCount != 0) {
                perTotal = (sumDailyReport.getTotal() - lastMonthTotalCount) / lastMonthTotalCount * 100;
            }
            perDailyReport.setTotal(perTotal);
        }
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getStartDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = customerOrderPlanDailyRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 客户每日派单报表导出
     *
     * @return
     */

    public SXSSFWorkbook customerOrderPlanDailyRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTCustomerOrderPlanDailyEntity> list = getCustomerOrderPlanDailyRptData(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 6));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算类型");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编码");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 6, days + 5));
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日下单情况(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 6, days + 6));
            ExportExcel.createCell(headFirstRow, days + 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex + 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCustomerOrderPlanDailyEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    if(dataRowIndex < rowsCount -3){
                        dataCell.setCellValue(dataRowIndex + 1);
                    }

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
                        ;
                    } else {
                        strTotalCount = String.format("%.0f", totalCount);
                    }
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                }
            }


        } catch (Exception e) {
            log.error("【CustomerOrderPlanDailyRptService.customerOrderPlanDailyRptExport】客户每日下单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


    /**
     * 获取客户每日派单报表转到每日工单图表

     * @return
     */
    public Map<String, Object> turnToCustomerOrderPlanDailyChartInformation(RPTCustomerOrderPlanDailySearch search) {
        Date startDate = new Date(search.getStartDate());
        int days = DateUtils.getDaysOfMonth(startDate);
        List<Integer> daySums = new ArrayList<>();

        List<String> createDates = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();

        List<RPTCustomerOrderPlanDailyEntity> copList = getCustomerOrderPlanDailyRptData(search);

        RPTBaseDailyEntity entity = copList.get(copList.size() - 3);
        copList.remove(copList.size()-1);
        copList.remove(copList.size()-1);
        copList.remove(copList.size()-1);
        copList =  copList.stream().sorted(Comparator.comparing(RPTCustomerOrderPlanDailyEntity::getTotal).reversed()).collect(Collectors.toList());
        int size = copList.size();
        copList = copList.subList(0,size>10?10:size);
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        if (entity != null) {
            Class itemClass = entity.getClass();
            for (int i = 0; i < days; i++) {
                String Date = DateUtils.formatDate(DateUtils.addDays(startDate, i), "dd");
                createDates.add(Date);
                String strGetMethodName = "getD" + (i + 1);
                try {
                    Method itemGetMethod = itemClass.getMethod(strGetMethodName);
                    Object itemGetD = itemGetMethod.invoke(entity);
                    Double dSum = StringUtils.toDouble(itemGetD);
                    int daySum = (int) Math.floor(dSum);

                    daySums.add(daySum);

                } catch (Exception ex) {
                    log.error("CustomerOrderPlanDailyRptService.turnToCustomerOrderPlanDayChartInformation",ex);
                }
            }
        }
        map.put("list",copList);
        map.put("createDates", createDates);
        map.put("daySums", daySums);
        return map;
    }
}
