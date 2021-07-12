package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTCustomerRevenueEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerRevenueSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.mapper.CustomerRevenueRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
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

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerRevenueRptService extends RptBaseService {
    @Resource
    private CustomerRevenueRptMapper customerRevenueRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    public List<RPTCustomerRevenueEntity> getCustomerRevenueList(RPTCustomerRevenueSearch search) {
        Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        Date startDate = DateUtils.getStartDayOfMonth(queryDate);
        Date endDate = DateUtils.getLastDayOfMonth(queryDate);
        Long startDt = startDate.getTime();
        Long endDt = endDate.getTime();
        int systemId = RptCommonUtils.getSystemId();
        Long customerId = search.getCustomerId();
        List<RPTCustomerRevenueEntity> list = customerRevenueRptMapper.getCustomerRevenueList(systemId, customerId,startDt, endDt,search.getProductCategoryIds(),quarter);

        list = list.stream().filter(i -> i.getCustomerId() != 0).collect(Collectors.toList());

        Set<Long> customerIds = list.stream().map(RPTCustomerRevenueEntity::getCustomerId).collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "name","salesId"};
        Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));

        RPTCustomer customer;
        RPTUser sales;
        int finishOrder;
        double receivableCharge;
        double payableCharge;
        double noPayableCharge;
        int finishOrderSum = 0;
        double receivableCharged = 0.0;
        double payableCharged = 0.0;
        double noPayableCharged = 0.0;
        double orderGrossProfit = 0.0;
        double noOrderGrossProfit = 0.0;
        Set<Long> salesIds = new HashSet<>();
        for (RPTCustomerRevenueEntity entity : list) {
            if (customerMap != null) {
                customer = customerMap.get(entity.getCustomerId());
                if (customer != null) {
                    entity.setCustomerName(customer.getName());
                    sales = customer.getSales();
                    if (sales != null) {
                        entity.setSalesId(sales.getId());
                        salesIds.add(entity.getSalesId());
                    }
                }
            }

            finishOrder = entity.getFinishOrder();
            receivableCharge = entity.getReceivableCharge();
            payableCharge = entity.getPayableCharge();
            noPayableCharge = entity.getNoPayableCharge();
            entity.setOrderGrossProfit(receivableCharge - payableCharge);
            entity.setNoOrderGrossProfit(receivableCharge-noPayableCharge);
            if (finishOrder != 0) {
                entity.setEverySingleGrossProfit(entity.getOrderGrossProfit() / finishOrder);
                entity.setNoEverySingleGrossProfit(entity.getNoOrderGrossProfit()/finishOrder);
            }

            finishOrderSum += entity.getFinishOrder();
            receivableCharged += entity.getReceivableCharge();
            payableCharged += entity.getPayableCharge();
            noPayableCharged += entity.getNoPayableCharge();
            orderGrossProfit += entity.getOrderGrossProfit();
            noOrderGrossProfit += entity.getNoOrderGrossProfit();


        }
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));
        for(RPTCustomerRevenueEntity entity : list){
            entity.setSalesName(userNameMap.get(entity.getSalesId()));
        }
        list = list.stream().sorted(Comparator.comparing(RPTCustomerRevenueEntity::getFinishOrder).reversed()).collect(Collectors.toList());
        //合计
        RPTCustomerRevenueEntity entity = new RPTCustomerRevenueEntity();
        entity.setFinishOrder(finishOrderSum);
        entity.setReceivableCharge(receivableCharged);
        entity.setPayableCharge(payableCharged);
        entity.setNoPayableCharge(noPayableCharged);
        entity.setOrderGrossProfit(orderGrossProfit);
        entity.setNoOrderGrossProfit(noOrderGrossProfit);
        if (orderGrossProfit != 0 && finishOrderSum != 0) {
            entity.setEverySingleGrossProfit(orderGrossProfit / finishOrderSum);
        }
        if (noOrderGrossProfit != 0 && finishOrderSum != 0) {
            entity.setNoEverySingleGrossProfit(noOrderGrossProfit / finishOrderSum);
        }
        list.add(entity);
        return list;
    }
    /**
     * 转到客户营收排名
     * @return
     */
    public Map<String, Object> turnToChartInformation(RPTCustomerRevenueSearch search) {
        Map<String, Object> map = new HashMap<>();
        List<RPTCustomerRevenueEntity> list = getCustomerRevenueList(search);

        List<RPTCustomerRevenueEntity> totalReceivables = new ArrayList<>();
        List<RPTCustomerRevenueEntity> finishOrders = new ArrayList<>();
        List<RPTCustomerRevenueEntity> orderGrossProfits = new ArrayList<>();
        List<RPTCustomerRevenueEntity> everySingleGrossProfits = new ArrayList<>();
        list.remove(list.size() - 1);

        totalReceivables = list.stream().sorted(Comparator.comparing(RPTCustomerRevenueEntity::getReceivableCharge).reversed()).collect(Collectors.toList());
        finishOrders = list.stream().sorted(Comparator.comparing(RPTCustomerRevenueEntity::getFinishOrder).reversed()).collect(Collectors.toList());
        orderGrossProfits = list.stream().sorted(Comparator.comparing(RPTCustomerRevenueEntity::getOrderGrossProfit).reversed()).collect(Collectors.toList());
        everySingleGrossProfits = list.stream().sorted(Comparator.comparing(RPTCustomerRevenueEntity::getEverySingleGrossProfit).reversed()).collect(Collectors.toList());

        int totalReceivableSize = totalReceivables.size();
        int finishOrderSize = finishOrders.size();
        int orderGrossProfitSize = orderGrossProfits.size();
        int everySingleGrossProfitSize = everySingleGrossProfits.size();

        totalReceivables = totalReceivables.subList(0, totalReceivableSize > 10 ? 10 : totalReceivableSize);
        finishOrders = finishOrders.subList(0, finishOrderSize > 10 ? 10 : finishOrderSize);
        orderGrossProfits = orderGrossProfits.subList(0, orderGrossProfitSize > 10 ? 10 : orderGrossProfitSize);
        everySingleGrossProfits = everySingleGrossProfits.subList(0, everySingleGrossProfitSize > 10 ? 10 : everySingleGrossProfitSize);

        map.put("totalReceivables", totalReceivables);
        map.put("finishOrders", finishOrders);
        map.put("orderGrossProfits", orderGrossProfits);
        map.put("everySingleGrossProfits", everySingleGrossProfits);

        return map;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerRevenueSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerRevenueSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getSelectedYear() != null && searchCondition.getSelectedMonth() != null) {
            Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
            Date startDate = DateUtils.getStartDayOfMonth(queryDate);
            Date endDate = DateUtils.getLastDayOfMonth(queryDate);
            String quarter = QuarterUtils.getSeasonQuarter(queryDate);
            int systemId = RptCommonUtils.getSystemId();
            Long startDt = startDate.getTime();
            Long endDt = endDate.getTime();
            Integer rowCount = customerRevenueRptMapper.hasReportData(systemId, startDt, endDt,searchCondition.getProductCategoryIds(),quarter);
            result = rowCount > 0;
        }
        return result;
    }

    private List<RPTCustomerRevenueEntity> getCustomerRevenueData(Date date) {
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date startDate = DateUtils.getStartOfDay(date);
        Date endDate = DateUtils.getEndOfDay(date);
        List<RPTCustomerRevenueEntity> list = new ArrayList<>();
        int systemId = RptCommonUtils.getSystemId();
        //获取完成工单数
        List<RPTCustomerRevenueEntity> finishOrderData = customerRevenueRptMapper.getFinishOrderData(startDate, endDate, quarter);
        Map<String, Integer> finishOrderMap = finishOrderData.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), RPTCustomerRevenueEntity::getFinishOrder));

        //获取应收合计
        List<RPTCustomerRevenueEntity> receivableChargeData = customerRevenueRptMapper.getReceivableCharge(startDate, endDate, quarter);
        Map<String, Double> receivableChargeMap = receivableChargeData.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), RPTCustomerRevenueEntity::getReceivableCharge));
        //获取应付1
        List<RPTCustomerRevenueEntity> payableChargeAData = customerRevenueRptMapper.getPayableChargeA(startDate, endDate, quarter);
        Map<String, Double> payableChargeAMap = payableChargeAData.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), RPTCustomerRevenueEntity::getPayableCharge));

        //获取应付2
        List<RPTCustomerRevenueEntity> payableChargeBData = customerRevenueRptMapper.getPayableChargeB(startDate, endDate, quarter);
        Map<String, RPTCustomerRevenueEntity> payableChargeBMap = payableChargeBData.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity(), (key1, key2) -> key2));

        Set<String> keys = new HashSet<>();
        keys.addAll(finishOrderMap.keySet());
        keys.addAll(receivableChargeMap.keySet());
        keys.addAll(payableChargeAMap.keySet());
        keys.addAll(payableChargeBMap.keySet());
        Integer finishOrder;
        Double receivableCharge;
        Double payableChargeA;
        Double payableChargeB;
        long customerId;
        long productCategoryId;
        RPTCustomerRevenueEntity entity;
        for (String key : keys) {
            entity = new RPTCustomerRevenueEntity();
            payableChargeA = 0.0;
            payableChargeB = 0.0;
            String[] split = key.split(":");
            customerId = Long.valueOf(split[0]);
            productCategoryId = Long.valueOf(split[1]);
            entity.setCustomerId(customerId);
            entity.setProductCategoryId(productCategoryId);
            entity.setSystemId(systemId);
            finishOrder = finishOrderMap.get(key);
            entity.setCreateDate(startDate.getTime());
            if (finishOrder != null) {
                entity.setFinishOrder(finishOrder);
            }
            receivableCharge = receivableChargeMap.get(key);
            if (receivableCharge != null) {
                entity.setReceivableCharge(receivableCharge);
            }
            if(payableChargeAMap.get(key) != null){
                payableChargeA = payableChargeAMap.get(key);
            }
            if(payableChargeBMap.get(key) != null){
                RPTCustomerRevenueEntity payableB = payableChargeBMap.get(key);
                payableChargeB = payableB.getPayableCharge();
                entity.setEngineerDeposit(payableB.getEngineerDeposit());
                entity.setEngineerInsuranceCharge(payableB.getEngineerInsuranceCharge());
            }

            entity.setPayableCharge(payableChargeA + payableChargeB);
            entity.setQuarter(quarter);

            list.add(entity);
        }

        return list;
    }

    public void saveCustomerRevenueToRptDB(Date date) {
        List<RPTCustomerRevenueEntity> list = getCustomerRevenueData(date);

        List<List<RPTCustomerRevenueEntity>> parts = Lists.partition(list, 100);


        for (List<RPTCustomerRevenueEntity> part : parts) {
            customerRevenueRptMapper.insertCustomerRevenueData(part);
            try {
                TimeUnit.MILLISECONDS.sleep(1000);
            } catch (InterruptedException e) {
                log.error("【CustomerRevenueRptService.saveCustomerRevenueToRptDB】客户营收数据写入中间表失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }
        }

    }

    /**
     * 删除中间表中指定日期的数据
     */
    private void deleteCustomerOrderTimeFromRptDB(Date date) {
        if (date != null) {
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int systemId = RptCommonUtils.getSystemId();
            customerRevenueRptMapper.deleteCustomerRevenueFromRptDB(systemId, startDate.getTime(), endDate.getTime(), quarter);
        }
    }

    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveCustomerRevenueToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteCustomerOrderTimeFromRptDB(beginDate);
                            saveCustomerRevenueToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteCustomerOrderTimeFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CustomerRevenueRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 报表导出
     */
    public SXSSFWorkbook customerRevenueExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerRevenueSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerRevenueSearch.class);
            List<RPTCustomerRevenueEntity> list = getCustomerRevenueList(searchCondition);
            NumberFormat numberFormat = NumberFormat.getInstance();
            numberFormat.setMaximumFractionDigits(2);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 4, 5));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 7));
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单毛利(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 8, 8));
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成单量(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 10));
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每单毛利(元/单)");


            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未扣除其他项");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣除其他项");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未扣除其他项");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣除其他项");
            ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未扣除其他项");
            ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣除其他项");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            if (list != null && list.size() > 0) {

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCustomerRevenueEntity entity = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;
                    if (dataRowIndex == list.size() - 1) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 2));
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getReceivableCharge()));
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getPayableCharge()));
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoPayableCharge()));
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getOrderGrossProfit()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoOrderGrossProfit()));
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getFinishOrder());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getEverySingleGrossProfit()));
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoEverySingleGrossProfit()));
                    } else {
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCustomerName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getSalesName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getReceivableCharge()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getPayableCharge()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoPayableCharge()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getOrderGrossProfit()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoOrderGrossProfit()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getFinishOrder());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getEverySingleGrossProfit()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, numberFormat.format(entity.getNoEverySingleGrossProfit()));
                    }
                }
            }

        } catch (Exception e) {
            log.error("【CustomerRevenueRptService.customerRevenueExport】客户营收报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
