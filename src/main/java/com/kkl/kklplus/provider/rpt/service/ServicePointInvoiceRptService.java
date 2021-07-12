package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointInvoiceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointInvoiceSearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointInvoiceRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
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
public class ServicePointInvoiceRptService extends  RptBaseService {

    @Autowired
    private ServicePointInvoiceRptMapper servicePointInvoiceRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    public Page<RPTServicePointInvoiceEntity> getServicePointInvoiceRptPage(RPTServicePointInvoiceSearch search){
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }
        Date endDate;
        Date beginDate;
        if(search.getEndDate()== 0L) {
           endDate = null;
        }else {
            endDate = new Date(search.getEndDate());

        }

        if(search.getBeginDate() == 0L){
            beginDate = null;
        }else {
            beginDate  = new Date(search.getBeginDate());
        }
        Date  beginInvoiceDate  =  new Date(search.getBeginInvoiceDate());
        Date  endInvoiceDate = new Date(search.getEndInvoiceDate());
        Page<RPTServicePointInvoiceEntity> page = servicePointInvoiceRptMapper.getServicePointInvoicePage(search.getServicePointId(),
                search.getWithdrawNo(), search.getPaymentType(), search.getBank(), beginDate, endDate, beginInvoiceDate, endInvoiceDate,search.getQuarters(),search.getStatus(),search.getPageNo(),search.getPageSize());


        // 获取网点的数据
        List<Long> servicePointIds = page.stream().map(RPTServicePointInvoiceEntity::getServicePointId).distinct().collect(Collectors.toList());

        String[] fieldsArray = new String[]{"id","servicePointNo","name","bankIssue"};
        Map<Long,MDServicePointViewModel> mdServicePointViewModelMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray),null);
//        List<Long> engineerIds = mdServicePointViewModelMap.entrySet().stream().map(i -> i.getValue().getPrimaryId()).distinct().collect(Collectors.toList());
//        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));

        Map<String, RPTDict> bankIssueTypeMap = MSDictUtils.getDictMap("BankIssueType");
        RPTDict bankIssue;
        for (RPTServicePointInvoiceEntity rptServicePointInvoiceEntity : page){
            MDServicePointViewModel servicePointVM = mdServicePointViewModelMap.get(rptServicePointInvoiceEntity.getServicePointId());
            if (servicePointVM != null) {
                rptServicePointInvoiceEntity.setServicePointNo(servicePointVM.getServicePointNo());
                rptServicePointInvoiceEntity.setServicePointName(servicePointVM.getName());
                rptServicePointInvoiceEntity.setContactInfo1(servicePointVM.getContactInfo1());
                rptServicePointInvoiceEntity.setContactInfo2(servicePointVM.getContactInfo2());
//                RPTEngineer engineer = engineerMap.get(servicePointVM.getPrimaryId());
                if(rptServicePointInvoiceEntity.getStatus() == 30){
                    bankIssue  = bankIssueTypeMap.get(String.valueOf(servicePointVM.getBankIssue()));
                    if (bankIssue != null && bankIssue.getLabel() != null) {
                        rptServicePointInvoiceEntity.setRemarks(bankIssue.getLabel());
                    }
                }
//                rptServicePointInvoiceEntity.setEngineerName(engineer == null ? "" : engineer.getName());
            }

        }


        //切换为微服务
        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap("bankType");
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");
        for (RPTServicePointInvoiceEntity item : page) {
            //设置银行名称
            RPTDict itemBank = item.getBank();
            if (itemBank != null && itemBank.getValue() != null) {
                item.setBank(bankMap.get(itemBank.getValue()));
            }
            //设置支付类型名称
            RPTDict itemPaymentType = item.getPaymentType();
            if (itemPaymentType != null && itemPaymentType.getValue() != null) {
                item.setPaymentType(paymentTypeMap.get(itemPaymentType.getValue()));
            }
        }

        return page;
    }

    public List<RPTServicePointInvoiceEntity> getServicePointInvoiceRptList(RPTServicePointInvoiceSearch search){
        Date endDate;
        Date beginDate;
        if(search.getEndDate()== 0L) {
            endDate = null;
        }else {
            endDate = new Date(search.getEndDate());

        }

        if(search.getBeginDate() == 0L){
            beginDate = null;
        }else {
            beginDate  = new Date(search.getBeginDate());
        }
        Date  beginInvoiceDate  =  new Date(search.getBeginInvoiceDate());
        Date  endInvoiceDate = new Date(search.getEndInvoiceDate());
        List<RPTServicePointInvoiceEntity> list = servicePointInvoiceRptMapper.getServicePointInvoiceList(search.getServicePointId(),
                search.getWithdrawNo(), search.getPaymentType(), search.getBank(), beginDate, endDate, beginInvoiceDate, endInvoiceDate,search.getQuarters(),search.getStatus());


        // 获取网点的数据
        List<Long> servicePointIds = list.stream().map(RPTServicePointInvoiceEntity::getServicePointId).distinct().collect(Collectors.toList());

        String[] fieldsArray = new String[]{"id","servicePointNo","name","bankIssue"};
        Map<Long,MDServicePointViewModel> mdServicePointViewModelMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray),null);
//        List<Long> engineerIds = mdServicePointViewModelMap.entrySet().stream().map(i -> i.getValue().getPrimaryId()).distinct().collect(Collectors.toList());
//        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));


        Map<String, RPTDict> bankIssueTypeMap = MSDictUtils.getDictMap("BankIssueType");
        RPTDict bankIssue;
        for (RPTServicePointInvoiceEntity rptServicePointInvoiceEntity : list){
            MDServicePointViewModel servicePointVM = mdServicePointViewModelMap.get(rptServicePointInvoiceEntity.getServicePointId());
            if (servicePointVM != null) {
                rptServicePointInvoiceEntity.setServicePointNo(servicePointVM.getServicePointNo());
                rptServicePointInvoiceEntity.setServicePointName(servicePointVM.getName());
                rptServicePointInvoiceEntity.setContactInfo1(servicePointVM.getContactInfo1());
                rptServicePointInvoiceEntity.setContactInfo2(servicePointVM.getContactInfo2());
//                RPTEngineer engineer = engineerMap.get(servicePointVM.getPrimaryId());
                if(rptServicePointInvoiceEntity.getStatus() == 30){
                    bankIssue  = bankIssueTypeMap.get(String.valueOf(servicePointVM.getBankIssue()));
                    if (bankIssue != null && bankIssue.getLabel() != null) {
                        rptServicePointInvoiceEntity.setRemarks(bankIssue.getLabel());
                    }
                }
//                rptServicePointInvoiceEntity.setEngineerName(engineer == null ? "" : engineer.getName());
            }
        }


        //切换为微服务
        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap("bankType");
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");
        for (RPTServicePointInvoiceEntity item : list) {
            //设置银行名称
            RPTDict itemBank = item.getBank();
            if (itemBank != null && itemBank.getValue() != null) {
                item.setBank(bankMap.get(itemBank.getValue()));
            }
            //设置支付类型名称
            RPTDict itemPaymentType = item.getPaymentType();
            if (itemPaymentType != null && itemPaymentType.getValue() != null) {
                item.setPaymentType(paymentTypeMap.get(itemPaymentType.getValue()));
            }
        }

        return list;
    }
    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointInvoiceSearch search = redisGsonService.fromJson(searchConditionJson, RPTServicePointInvoiceSearch.class);
        Date endDate;
        Date beginDate;
        if(search.getEndDate()== 0L) {
            endDate = null;
        }else {
            endDate = new Date(search.getEndDate());
        }

        if(search.getBeginDate() == 0L){
            beginDate = null;
        }else {
            beginDate  = new Date(search.getBeginDate());
        }
        Date  beginInvoiceDate  =  new Date((search.getBeginInvoiceDate()));
        Date  endInvoiceDate = new Date(search.getEndInvoiceDate());
        if (beginInvoiceDate != null) {
            Integer rowCount = servicePointInvoiceRptMapper.hasReportData(search.getServicePointId(),
                    search.getWithdrawNo(), search.getPaymentType(), search.getBank(), beginDate, endDate, beginInvoiceDate, endInvoiceDate,search.getQuarters(),search.getStatus());
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook servicePointInvoiceRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointInvoiceSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointInvoiceSearch.class);
            List<RPTServicePointInvoiceEntity> list = getServicePointInvoiceRptList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");

            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户银行");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户人");
            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "银行账号");

            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结账单号");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "请款日期");
            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款日期");

            ExportExcel.createCell(headRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款金额");
            ExportExcel.createCell(headRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台服务费");
            ExportExcel.createCell(headRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款描述");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if (list != null && list.size() > 0) {
                double totalPayAmount = 0.0;
                double totalPlatformFee = 0.0;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTServicePointInvoiceEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointName());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBank() == null ? "" : rowData.getBank().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankOwner());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getWithdrawNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getPayDate(), "yyyy-MM-dd HH:mm:ss"));

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPayAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPlatformFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRemarks());

                    totalPayAmount = totalPayAmount + rowData.getPayAmount();
                    totalPlatformFee = totalPlatformFee + rowData.getPlatformFee();
                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 9));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPayAmount);
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPlatformFee);
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
            }



        } catch (Exception e) {
            log.error("【ServicePointInvoiceRptService.servicePointInvoiceRptExport】网点付款清单写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
