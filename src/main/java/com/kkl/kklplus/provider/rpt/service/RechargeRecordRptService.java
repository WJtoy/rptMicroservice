package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTRechargeRecordEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.RechargeRecordRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
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
public class RechargeRecordRptService extends RptBaseService {

    @Autowired
    private RechargeRecordRptMapper rechargeRecordRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    public Page<RPTRechargeRecordEntity> getRechargeRecordByPage(RPTCustomerOrderPlanDailySearch search) {
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        Date startDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        Page<RPTRechargeRecordEntity> rechargeRptEntities = rechargeRecordRptMapper.rechargeRecordPage(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters(), search.getPageNo(), search.getPageSize());
        Set<Long> customerIds = rechargeRptEntities.stream().map(RPTRechargeRecordEntity::getCustomerId).distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Map<String, RPTDict> customerActionType = MSDictUtils.getDictMap("CustomerActionType");
        for (RPTRechargeRecordEntity entity : rechargeRptEntities) {
            RPTCustomer customer = customerMap.get(entity.getCustomerId());
            if (customer != null) {
                entity.setCustomerName(customer.getName());
            }
            if (customerActionType != null) {
                entity.setActionType(customerActionType.get(entity.getActionType().getValue()));
            }
        }
        return rechargeRptEntities;
    }

    public List<RPTRechargeRecordEntity> getRechargeRecordByList(RPTCustomerOrderPlanDailySearch search) {

        Date startDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());
        List<RPTRechargeRecordEntity> rechargeRptEntities = rechargeRecordRptMapper.rechargeRecordList(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters());
        Set<Long> customerIds = rechargeRptEntities.stream().map(RPTRechargeRecordEntity::getCustomerId).distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Map<String, RPTDict> customerActionType = MSDictUtils.getDictMap("CustomerActionType");
        for (RPTRechargeRecordEntity entity : rechargeRptEntities) {
            RPTCustomer customer = customerMap.get(entity.getCustomerId());
            if (customer != null) {
                entity.setCustomerName(customer.getName());
            }
            if (customerActionType != null) {
                entity.setActionType(customerActionType.get(entity.getActionType().getValue()));
            }
        }
        return rechargeRptEntities;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch search = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        if (search.getStartDate() != null && search.getEndDate() != null) {
            Date startDate = new Date(search.getStartDate());
            Date endDate = new Date(search.getEndDate());
            Long rowCount = rechargeRecordRptMapper.hasReportData(search.getCustomerId(), startDate, endDate, search.getPaymentType(), search.getQuarters());
           if(rowCount ==null){
               result = false;
           }
        }
        return result;
    }

    public SXSSFWorkbook rechargeRecordRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTRechargeRecordEntity> rechargeRecordList = getRechargeRecordByList(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 8));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headerSecondRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值类型");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "创建时间");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值时间");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "充值金额");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "收入方式");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "描述");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "相关单号");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================

            if (rechargeRecordList != null) {
                int rowNumber = 0;
                Double actionType10Sum = 0.0;
                Double actionType20Sum = 0.0;
                Double actionType90Sum = 0.0;
                int line = 0;
                int fi = 0;
                int off = 0;
                for (RPTRechargeRecordEntity orderMaster : rechargeRecordList) {
                    rowNumber++;
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomerName()));
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getActionType().getLabel() != null ? orderMaster.getActionType().getLabel() : ""));
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getUpdateDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getAmount()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getActionType().getValue().equals("10") ? "银行" : orderMaster.getActionType().getValue().equals("20") ? "现金" : ""));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getRemarks()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCurrencyNo()));
                    if (orderMaster.getActionType().getValue().equals("10")) {
                        actionType10Sum = actionType10Sum + orderMaster.getAmount();
                        line = line + 1;
                    }
                    if (orderMaster.getActionType().getValue().equals("20")) {
                        actionType20Sum = actionType20Sum + orderMaster.getAmount();
                        fi = fi + 1;
                    }
                    if (orderMaster.getActionType().getValue().equals("90")) {
                        actionType90Sum = actionType90Sum + orderMaster.getAmount();
                        off = off + 1;
                    }
                }


                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "在线充值：" + line + "笔，合计" + actionType10Sum + "元,财务充值：" + fi + "笔，合计" + actionType20Sum + "元,线下充值：" + off + "笔，合计" + actionType90Sum + "元，累计充值：" + (actionType10Sum + actionType20Sum + actionType90Sum) + "元");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 2, 8));

            }

        } catch (Exception e) {
            log.error("【RechargeRecordRptService.rechargeRecordRptExport】充值明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
