package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTCustomerRechargeSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.provider.rpt.entity.FIEnums;
import com.kkl.kklplus.provider.rpt.mapper.DepositRechargeSummaryMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
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

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class DepositRechargeSummaryRptService extends RptBaseService {

    @Autowired
    private DepositRechargeSummaryMapper depositRechargeSummaryMapper;

    @Autowired
    private MSCustomerService customerService;

    @Autowired
    private MSServicePointService msServicePointService;

    public List<RPTCustomerRechargeSummaryEntity> getDepositRechargeSummaryList(RPTCustomerOrderPlanDailySearch search) {
        Date startDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        List<RPTCustomerRechargeSummaryEntity> list = new ArrayList<>();
        List<RPTCustomerRechargeSummaryEntity> onlineList = depositRechargeSummaryMapper.getOnlineRechargeSummaryList(search.getCustomerId(), startDate, endDate, search.getQuarter());

        List<RPTCustomerRechargeSummaryEntity> orderDeductList = depositRechargeSummaryMapper.getOrderCompleteDeductSummaryList(search.getCustomerId(), startDate, endDate, search.getQuarter());


        Date preStartDate = DateUtils.addMonth(startDate, -1);
        Date preEndDate = DateUtils.getLastDayOfMonth(preStartDate);
        String preQuarter = DateUtils.getQuarter(preStartDate);
        List<RPTCustomerRechargeSummaryEntity> preOnlineList = depositRechargeSummaryMapper.getOnlineRechargeSummaryList(search.getCustomerId(), preStartDate, preEndDate, preQuarter);
        List<RPTCustomerRechargeSummaryEntity> preOrderDeductList = depositRechargeSummaryMapper.getOrderCompleteDeductSummaryList(search.getCustomerId(), preStartDate, preEndDate, preQuarter);


        Set<Long> servicePointIdSet = Sets.newHashSet();
        Map<Long, RPTCustomerRechargeSummaryEntity> onlineMap = Maps.newHashMap();
        Map<Long, RPTCustomerRechargeSummaryEntity> orderDeductMap = Maps.newHashMap();
        Map<Long, RPTCustomerRechargeSummaryEntity> preOnlineMap = Maps.newHashMap();
        Map<Long, RPTCustomerRechargeSummaryEntity> preOrderDeductMap = Maps.newHashMap();
        for (RPTCustomerRechargeSummaryEntity item : onlineList) {
            servicePointIdSet.add(item.getCustomerId());
            onlineMap.put(item.getCustomerId(), item);
        }
        for (RPTCustomerRechargeSummaryEntity item : orderDeductList) {
            servicePointIdSet.add(item.getCustomerId());
            orderDeductMap.put(item.getCustomerId(), item);
        }
        for (RPTCustomerRechargeSummaryEntity item : preOnlineList) {
            preOnlineMap.put(item.getCustomerId(), item);
        }
        for (RPTCustomerRechargeSummaryEntity item : preOrderDeductList) {
            preOrderDeductMap.put(item.getCustomerId(), item);
        }


        List<Long> servicePointIds = Lists.newArrayList(servicePointIdSet);
        String[] fieldsArray = new String[]{"id", "servicePointNo", "name",};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray), null);
        RPTCustomerRechargeSummaryEntity rptEntity;
        RPTCustomerRechargeSummaryEntity onlineEntity;
        RPTCustomerRechargeSummaryEntity orderDeductEntity;
        RPTCustomerRechargeSummaryEntity preOnlineEntity;
        RPTCustomerRechargeSummaryEntity preOrderDeductEntity;
        double engineerDeposit;
        double rechargeDeposit;
        for (Long servicePointId : servicePointIds) {
            rptEntity = new RPTCustomerRechargeSummaryEntity();

            onlineEntity = onlineMap.get(servicePointId);
            orderDeductEntity = orderDeductMap.get(servicePointId);
            preOnlineEntity = preOnlineMap.get(servicePointId);
            preOrderDeductEntity = preOrderDeductMap.get(servicePointId);
            rptEntity.setOnlineRechargeAmount(0.0);
            rptEntity.setFinancialRechargeAmount(0.0);
            rptEntity.setOfflineRechargeAmount(0.0);
            engineerDeposit = 0.0;
            rechargeDeposit = 0.0;
            if (onlineEntity != null) {
                rptEntity.setOnlineRechargeAmount(onlineEntity.getOnlineRechargeAmount());
            }
            if (orderDeductEntity != null) {
                rptEntity.setFinancialRechargeAmount(orderDeductEntity.getFinancialRechargeAmount());
            }
            if (preOnlineEntity != null) {
                rechargeDeposit = preOnlineEntity.getOnlineRechargeAmount();
            }
            if (preOrderDeductEntity != null) {
                engineerDeposit = preOrderDeductEntity.getFinancialRechargeAmount();
            }
            rptEntity.setBeforeBalance(rechargeDeposit + engineerDeposit);

            MDServicePointViewModel servicePoint = servicePointMap.get(servicePointId);
            if (servicePoint != null) {
                rptEntity.setCustomerName(servicePoint.getName());
                rptEntity.setCustomerCode(servicePoint.getServicePointNo());
            }
            list.add(rptEntity);
        }

        list = list.stream().sorted(Comparator.comparing(RPTCustomerRechargeSummaryEntity::getCustomerName)).collect(Collectors.toList());

        return list;

    }


    public Page<RPTRechargeRecordEntity> getDepositRechargeDetailsByPage(RPTCustomerOrderPlanDailySearch search) {
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }
        Date startDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        Page<RPTRechargeRecordEntity> rechargeRptEntities = depositRechargeSummaryMapper.getDepositRechargeDetailsPage(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters(), search.getPageNo(), search.getPageSize());
        Set<Long> servicePointIds = rechargeRptEntities.stream().map(RPTRechargeRecordEntity::getCustomerId).distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "servicePointNo", "name"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        FIEnums.DepositActionTypeENum depositActionTypeENum;
        for (RPTRechargeRecordEntity entity : rechargeRptEntities) {
            MDServicePointViewModel servicePointViewModel = servicePointMap.get(entity.getCustomerId());
            if (servicePointViewModel != null) {
                entity.setCustomerName(servicePointViewModel.getName());
                entity.setCustomerCode(servicePointViewModel.getServicePointNo());
            }
            if (entity.getActionType() != null && !entity.getActionType().getValue().equals("")) {
                depositActionTypeENum = FIEnums.DepositActionTypeENum.fromValue(Integer.parseInt(entity.getActionType().getValue()));
                assert depositActionTypeENum != null;
                entity.getActionType().setLabel(depositActionTypeENum.getName());
            }
        }
        return rechargeRptEntities;
    }

    public List<RPTRechargeRecordEntity> getDepositRechargeDetailsByList(RPTCustomerOrderPlanDailySearch search) {

        Date startDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        List<RPTRechargeRecordEntity> rechargeRptEntities = depositRechargeSummaryMapper.getDepositRechargeDetailsList(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters());
        Set<Long> servicePointIds = rechargeRptEntities.stream().map(RPTRechargeRecordEntity::getCustomerId).distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "servicePointNo", "name"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        FIEnums.DepositActionTypeENum depositActionTypeENum;
        for (RPTRechargeRecordEntity entity : rechargeRptEntities) {
            MDServicePointViewModel servicePointViewModel = servicePointMap.get(entity.getCustomerId());
            if (servicePointViewModel != null) {
                entity.setCustomerName(servicePointViewModel.getName());
                entity.setCustomerCode(servicePointViewModel.getServicePointNo());
            }
            if (entity.getActionType() != null && !entity.getActionType().getValue().equals("")) {
                depositActionTypeENum = FIEnums.DepositActionTypeENum.fromValue(Integer.parseInt(entity.getActionType().getValue()));
                assert depositActionTypeENum != null;
                entity.getActionType().setLabel(depositActionTypeENum.getName());
            }
        }
        return rechargeRptEntities;
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
        Integer rowCount = depositRechargeSummaryMapper.hasReportData(searchCondition.getCustomerId(), beginDate, endDate, searchCondition.getQuarter());
        result = rowCount > 0;
        return result;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasDepositReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch search = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        if (search.getStartDate() != null && search.getEndDate() != null) {
            Date startDate = new Date(search.getStartDate());
            Date endDate = new Date(search.getEndDate());
            Long rowCount = depositRechargeSummaryMapper.hasDepositReportData(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters());
            if (rowCount == null) {
                result = false;
            }
        }
        return result;
    }

    public SXSSFWorkbook depositRechargeDetailsExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTRechargeRecordEntity> rechargeRecordList = getDepositRechargeDetailsByList(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 9));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headerSecondRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值类型");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "创建时间");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值时间");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值金额");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "收入方式");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "描述");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "相关单号");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================

            if (rechargeRecordList != null) {
                int rowNumber = 0;
                Double actionType10Sum = 0.0;
                Double actionType20Sum = 0.0;
                int line = 0;
                int fi = 0;
                for (RPTRechargeRecordEntity orderMaster : rechargeRecordList) {
                    rowNumber++;
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomerCode()));
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomerName()));
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getActionType().getLabel() != null ? orderMaster.getActionType().getLabel() : ""));
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getUpdateDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getAmount()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getActionType().getValue().equals("10") ? "银行" : orderMaster.getActionType().getValue().equals("20") ? "现金" : ""));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getRemarks()));
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCurrencyNo()));
                    if (orderMaster.getActionType().getValue().equals("10")) {
                        actionType10Sum = actionType10Sum + orderMaster.getAmount();
                        line = line + 1;
                    }
                    if (orderMaster.getActionType().getValue().equals("20")) {
                        actionType20Sum = actionType20Sum + orderMaster.getAmount();
                        fi = fi + 1;
                    }
                }


                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "线下充值：" + line + "笔，合计" + actionType10Sum + "元,订单完成扣款：" + fi + "笔，合计" + actionType20Sum + "元，累计充值：" + (actionType10Sum + actionType20Sum) + "元");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 2, 9));

            }

        } catch (Exception e) {
            log.error("【DepositRechargeSummaryRptService.depositRechargeDetailsExport】质保金充值明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


    public SXSSFWorkbook depositRechargeSummaryRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTCustomerRechargeSummaryEntity> list = getDepositRechargeSummaryList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月余额");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 4, 6));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月质保金充值金额");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "线下充值");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "订单完成扣款");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月余额");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if (list != null && list.size() > 0) {
                double onlineRechargeAmount = 0.0;
                double financialRechargeAmount = 0.0;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCustomerRechargeSummaryEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerCode());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBeforeBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOnlineRechargeAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getFinancialRechargeAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOnlineRechargeAmount() + rowData.getFinancialRechargeAmount());
                    onlineRechargeAmount = onlineRechargeAmount + rowData.getOnlineRechargeAmount();
                    financialRechargeAmount = financialRechargeAmount + rowData.getFinancialRechargeAmount();

                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 3));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, onlineRechargeAmount);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, financialRechargeAmount);
                ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, onlineRechargeAmount + financialRechargeAmount);

            }


        } catch (Exception e) {
            log.error("【DepositRechargeSummaryRptService.depositRechargeSummaryRptExport】客户质保金充值汇总写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


}
