package com.kkl.kklplus.provider.rpt.service;


import com.kkl.kklplus.entity.rpt.RPTCustomerReceivableSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointInvoiceSearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointPaymentSummaryRptMapper;
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
public class ServicePointPaymentSummaryRptService extends RptBaseService {

    @Autowired
    private ServicePointPaymentSummaryRptMapper servicePointPaymentSummaryRptMapper;

    public List<RPTServicePointInvoiceEntity> getServicePointPaymentSummaryList(RPTServicePointInvoiceSearch search) {
        Date beginDate = new Date(search.getBeginDate());
        Date endDate = new Date(search.getEndDate());
        List<RPTServicePointInvoiceEntity> list = servicePointPaymentSummaryRptMapper.getServicePointPaymentList(search.getPaymentType(), search.getBank(), beginDate, endDate, search.getQuarter());

        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap("bankType");
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");

        for (RPTServicePointInvoiceEntity item : list) {

            //设置银行名称
            RPTDict itemBank = bankMap.get(item.getBankOwner());
            if (itemBank != null && itemBank.getValue() != null) {
                item.setBank(itemBank);
            }
            //设置支付类型名称
            RPTDict itemPaymentType = paymentTypeMap.get(item.getBankNo());
            if (itemPaymentType != null && itemPaymentType.getValue() != null) {
                item.setPaymentType(itemPaymentType);
            }
        }


        list = list.stream().sorted(Comparator.comparing(RPTServicePointInvoiceEntity::getPayDate)).collect(Collectors.toList());

        return list;


    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointInvoiceSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointInvoiceSearch.class);
        Date beginDate = new Date(searchCondition.getBeginDate());
        Date endDate = new Date(searchCondition.getEndDate());
        if (new Date().getTime() < beginDate.getTime()) {
            return false;
        }
        Integer rowCount = servicePointPaymentSummaryRptMapper.hasReportData(searchCondition.getPaymentType(), searchCondition.getBank(), beginDate, endDate, searchCondition.getQuarter());
        result = rowCount > 0;
        return result;
    }




    public SXSSFWorkbook servicePointPaymentSummaryRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTServicePointInvoiceSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointInvoiceSearch.class);
            List<RPTServicePointInvoiceEntity> list = getServicePointPaymentSummaryList(searchCondition);
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
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款时间");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "银行");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "金额");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款金额");
            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台服务费");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if(list != null && list.size() > 0){
                double totalPayAmount = 0.0;
                double totalPlatformFee = 0.0;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTServicePointInvoiceEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getPayDate(), "yyyy-MM-dd"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBank() == null ? "" : rowData.getBank().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPayAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPlatformFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPayAmount()+rowData.getPlatformFee());
                    totalPayAmount = totalPayAmount + rowData.getPayAmount();
                    totalPlatformFee = totalPlatformFee + rowData.getPlatformFee();
                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 2));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPayAmount);
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPlatformFee);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,totalPayAmount+totalPlatformFee);

            }


        } catch (Exception e) {
            log.error("【ServicePointPaymentSummaryRptService.servicePointPaymentSummaryRptExport】网点付款汇总写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

}
