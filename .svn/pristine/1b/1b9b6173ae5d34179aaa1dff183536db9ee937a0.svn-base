package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTSalesPerfomanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.mapper.CustomerRechargeSummaryMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
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

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerRechargeSummaryRptService extends RptBaseService {

    @Autowired
    private CustomerRechargeSummaryMapper customerRechargeSummaryMapper;

    @Autowired
    private MSCustomerService customerService;

    public List<RPTCustomerRechargeSummaryEntity> getCustomerRechargeSummaryList(RPTCustomerOrderPlanDailySearch search) {
        Date startDate   = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        List<RPTCustomerRechargeSummaryEntity> list  = new ArrayList<>();
        List<RPTCustomerRechargeSummaryEntity> onlineList = customerRechargeSummaryMapper.getOnlineRechargeSummaryList(search.getCustomerId(),startDate,endDate,search.getQuarter());

        List<RPTCustomerRechargeSummaryEntity> financialList = customerRechargeSummaryMapper.getFinancialRechargeSummaryList(search.getCustomerId(),startDate,endDate,search.getQuarter());


        List<RPTCustomerRechargeSummaryEntity> rechargeList = customerRechargeSummaryMapper.getRechargeSummaryList(search.getCustomerId(),startDate,endDate,search.getQuarter());

        Set<Long> customerIdSet = Sets.newHashSet();
        Map<Long, RPTCustomerRechargeSummaryEntity> onlineMap = Maps.newHashMap();
        for (RPTCustomerRechargeSummaryEntity item : onlineList) {
            customerIdSet.add(item.getCustomerId());
            onlineMap.put(item.getCustomerId(), item);
        }

        Map<Long, RPTCustomerRechargeSummaryEntity> financialMap = Maps.newHashMap();
        for (RPTCustomerRechargeSummaryEntity item : financialList) {
            customerIdSet.add(item.getCustomerId());
            financialMap.put(item.getCustomerId(), item);
        }


        Map<Long, RPTCustomerRechargeSummaryEntity> rechargeMap = Maps.newHashMap();
        for (RPTCustomerRechargeSummaryEntity item : rechargeList) {
            customerIdSet.add(item.getCustomerId());
            rechargeMap.put(item.getCustomerId(), item);
        }

        Set<Long> userIds = Sets.newHashSet();
        List<Long> customerIds = Lists.newArrayList(customerIdSet);
        Map<Long, RPTCustomer> customerMap = customerService.getCustomerMapWithContractDate(customerIds);
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        userIds.addAll(salesIds);
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
        String userName;
        RPTCustomerRechargeSummaryEntity rptEntity;
        for(Long customerId : customerIds){
            rptEntity = new RPTCustomerRechargeSummaryEntity();

            RPTCustomerRechargeSummaryEntity onlineEntity  = onlineMap.get(customerId);

            RPTCustomerRechargeSummaryEntity financialEntity = financialMap.get(customerId);

            RPTCustomerRechargeSummaryEntity rechargeEntity = rechargeMap.get(customerId);

            rptEntity.setOnlineRechargeAmount(0.0);
            if(onlineEntity != null){
                rptEntity.setOnlineRechargeAmount(onlineEntity.getOnlineRechargeAmount());
            }

            rptEntity.setFinancialRechargeAmount(0.0);
            if(financialEntity != null){
                rptEntity.setFinancialRechargeAmount(financialEntity.getFinancialRechargeAmount());
            }

            rptEntity.setOfflineRechargeAmount(0.0);
            if(rechargeEntity != null){
                rptEntity.setOfflineRechargeAmount(rechargeEntity.getOfflineRechargeAmount());
            }

            RPTCustomer customer = customerMap.get(customerId);
            if(customer != null){
                rptEntity.setCustomerName(customer.getName());
                rptEntity.setCustomerCode(customer.getCode());
                userName = userNameMap.get(customer.getSales().getId());
                if (StringUtils.isNotBlank(userName)) {
                    rptEntity.setSalesName(userName);
                }
            }
            list.add(rptEntity);
        }

        list = list.stream().sorted(Comparator.comparing(RPTCustomerRechargeSummaryEntity::getCustomerName)).collect(Collectors.toList());

        return list;

    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        Date beginDate = new Date(searchCondition.getStartDate());
        Date endDate = new Date(searchCondition.getEndDate());
        if (new Date().getTime() < beginDate.getTime()) {
            return false;
        }
        Integer rowCount = customerRechargeSummaryMapper.hasReportData(searchCondition.getCustomerId(), beginDate, endDate, searchCondition.getQuarter());
        result = rowCount > 0;
        return result;
    }




    public SXSSFWorkbook customerRechargeSummaryRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTCustomerRechargeSummaryEntity> list = getCustomerRechargeSummaryList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 21));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 4, 7));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值金额");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "在线充值");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "财务充值");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "线下充值");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if(list != null && list.size() > 0){
                double onlineRechargeAmount = 0.0;
                double financialRechargeAmount = 0.0;
                double offlineRechargeAmount = 0.0;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCustomerRechargeSummaryEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerCode());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSalesName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOnlineRechargeAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getFinancialRechargeAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOfflineRechargeAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOnlineRechargeAmount()+rowData.getFinancialRechargeAmount()+rowData.getOfflineRechargeAmount());
                    onlineRechargeAmount = onlineRechargeAmount + rowData.getOnlineRechargeAmount();
                    financialRechargeAmount = financialRechargeAmount + rowData.getFinancialRechargeAmount();
                    offlineRechargeAmount = offlineRechargeAmount + rowData.getOfflineRechargeAmount();

                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 3));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, onlineRechargeAmount);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, financialRechargeAmount);
                ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, offlineRechargeAmount);
                ExportExcel.createCell(sumRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,onlineRechargeAmount+financialRechargeAmount+offlineRechargeAmount);

            }


        } catch (Exception e) {
            log.error("【CustomerRechargeSummaryRptService.customerRechargeSummaryRptExport】客户充值汇总写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }






}
