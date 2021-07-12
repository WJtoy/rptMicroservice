package com.kkl.kklplus.provider.rpt.customer.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.rpt.RPTCustomerFinanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerFinanceSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.mapper.CustomerFinanceRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
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
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCustomerFinanceRptService extends RptBaseService {
    @Resource
    private CustomerFinanceRptMapper customerFinanceRptMapper;
    @Autowired
    private MSCustomerService msCustomerService;

    public MSPage<RPTCustomerFinanceEntity> getCustomerFinanceData(RPTCustomerFinanceSearch search, Integer pageSize) {
        MSPage<RPTCustomerFinanceEntity> page = new MSPage<>();
        MSPage<RPTCustomer> entityPage = new MSPage<>();
        List<RPTCustomerFinanceEntity> list = new ArrayList<>();
        RPTCustomer customer = new RPTCustomer();
        if (search.getCustomerId() != null && search.getCustomerId() != 0) {
            customer.setId(search.getCustomerId());
        }
        if (StringUtils.isNotBlank(search.getCode())) {
            customer.setCode(search.getCode());
        }
        if(search.getSalesId() != null && search.getSalesId() != 0){
            customer.getSales().setId(search.getSalesId());
        }
        if (search.getPaymentType() != null && search.getPaymentType() != 0) {
            customer.getPaymentType().setValue(search.getPaymentType().toString());
        }
        if(search.getMerchandiserId() != null && search.getMerchandiserId() != 0){
            customer.getMerchandiser().setId(search.getMerchandiserId());
        }
        entityPage.setPageNo(search.getPageNo());
        entityPage.setPageSize(search.getPageSize());
        entityPage = msCustomerService.findCustomerListWithCodeNamePaySaleContract(entityPage, customer);
        List<RPTCustomer> customerList = entityPage.getList();
        Set<Long> customerIds;
        Set<Long> salesIds = new HashSet<>();
        RPTCustomerFinanceEntity finance;
        RPTCustomerFinanceEntity entity;
        customerIds = customerList.stream().map(RPTCustomer::getId).collect(Collectors.toSet());
        List<RPTCustomerFinanceEntity> customerFinance = new ArrayList<>();
        if (pageSize != null && pageSize > 100) {
            customerFinance = customerFinanceRptMapper.getAllCustomerFinance();
        } else {
            if(!customerIds.isEmpty()){
                customerFinance = customerFinanceRptMapper.getCustomerFinance(Lists.newArrayList(customerIds));
            }
        }

        Map<Long, RPTCustomerFinanceEntity> maps = customerFinance.stream().collect(Collectors.toMap(i -> i.getCustomer().getId(), Function.identity(), (key1, key2) -> key2));

        for (RPTCustomer rptCustomer : customerList) {
            entity = new RPTCustomerFinanceEntity();
            entity.setCustomer(rptCustomer);
            finance = maps.get(entity.getCustomer().getId());
            if (finance != null) {
                entity.setBalance(finance.getBalance());
                entity.setBlockAmount(finance.getBlockAmount());
                entity.setCredit(finance.getCredit());
            }
            salesIds.add(entity.getCustomer().getSales().getId());
            entity.setRemarks(rptCustomer.getRemarks());
            list.add(entity);
        }
        Long salesId;
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));
        for (RPTCustomerFinanceEntity rptEntity : list) {
            if (rptEntity.getCustomer().getSales() != null) {
                salesId = rptEntity.getCustomer().getSales().getId();
                if (salesId != null) {
                    rptEntity.getCustomer().getSales().setName(userNameMap.get(salesId));
                }
            }
        }
        page.setList(list);
        page.setPageSize(entityPage.getPageSize());
        page.setPageNo(entityPage.getPageNo());
        page.setRowCount(entityPage.getRowCount());
        return page;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerFinanceSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerFinanceSearch.class);
        RPTCustomer customer = new RPTCustomer();
        MSPage<RPTCustomer> entityPage = new MSPage<>();
        if (searchCondition.getCustomerId() != null && searchCondition.getCustomerId() != 0) {
            customer.setId(searchCondition.getCustomerId());
        }
        if (StringUtils.isNotBlank(searchCondition.getCode())) {
            customer.setCode(searchCondition.getCode());
        }
        if (searchCondition.getPaymentType() != null && searchCondition.getPaymentType() != 0) {
            customer.getPaymentType().setValue(searchCondition.getPaymentType().toString());
        }
        entityPage.setPageNo(searchCondition.getPageNo());
        entityPage.setPageSize(searchCondition.getPageSize());
        entityPage = msCustomerService.findCustomerListWithCodeNamePaySaleContract(entityPage, customer);
        Integer rowCount = entityPage.getList().size();
        result = rowCount > 0;

        return result;
    }

    public SXSSFWorkbook customerFinanceRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;

        try {
            RPTCustomerFinanceSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerFinanceSearch.class);
            int pageSize = searchCondition.getPageSize();
            MSPage<RPTCustomerFinanceEntity> page = getCustomerFinanceData(searchCondition, pageSize);

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
            Row headerRow = xSheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "代码");
            ExportExcel.createCell(headerRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "账户余额");
            ExportExcel.createCell(headerRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "冻结金额");
            ExportExcel.createCell(headerRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "可下单金额");
            ExportExcel.createCell(headerRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "信用额度");
            ExportExcel.createCell(headerRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单位");
            ExportExcel.createCell(headerRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");
            ExportExcel.createCell(headerRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "描述");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            Cell dataCell = null;
            if (page.getList() != null && page.getList().size() > 0) {

                double totalBalance = 0.0;
                double totalBlockAmount = 0.0;
                double totalAllowCreateOrderCharge = 0.0;
                double totalCredit = 0.0;

                int rowsCount = page.getList().size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCustomerFinanceEntity rowData = page.getList().get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomer().getCode());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomer().getName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            rowData.getBalance() == null ? 0 : rowData.getBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            rowData.getBlockAmount() == null ? 0 : rowData.getBlockAmount());
                    double allowCreateOrderCharge = (rowData.getBalance() == null ? 0 : rowData.getBalance()) -
                            (rowData.getBlockAmount() == null ? 0 : rowData.getBlockAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, allowCreateOrderCharge);

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            rowData.getCredit() == null ? 0 : rowData.getCredit());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "RMB");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            rowData.getCustomer().getSales() == null ? "" : rowData.getCustomer().getSales().getName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            rowData.getRemarks() == null ? "" : rowData.getRemarks());

                    totalBalance = totalBalance + (rowData.getBalance() == null ? 0 : rowData.getBalance());
                    totalBlockAmount = totalBlockAmount + (rowData.getBlockAmount() == null ? 0 : rowData.getBlockAmount());
                    totalAllowCreateOrderCharge = totalAllowCreateOrderCharge + allowCreateOrderCharge;
                    totalCredit = totalCredit + (rowData.getCredit() == null ? 0 : rowData.getCredit());
                }

                Row sumRow = xSheet.createRow(rowIndex);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 1));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                ExportExcel.createCell(sumRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBalance);
                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBlockAmount);
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalAllowCreateOrderCharge);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCredit);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 6, 8));
                ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

            }

        } catch (Exception e) {
            log.error("【CustomerFinanceRptService.customerFinanceRptExport】客户账户余额报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
