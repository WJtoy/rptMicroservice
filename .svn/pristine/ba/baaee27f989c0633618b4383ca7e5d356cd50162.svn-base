package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.GlobalMappingSalesSubFlagEnum;
import com.kkl.kklplus.entity.rpt.RPTCustomerReceivableSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.CustomerReceivableSummaryRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
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
import org.springframework.util.CollectionUtils;

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerReceivableSummaryRptService extends RptBaseService {

    @Autowired
    private CustomerReceivableSummaryRptMapper customerReceivableSummaryRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;


    public MSPage<RPTCustomerReceivableSummaryEntity> getCustomerReceivableSummaryMonth(RPTCustomerOrderPlanDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        MSPage<RPTCustomer> entityPage = new MSPage<>();
        MSPage<RPTCustomerReceivableSummaryEntity> page = new MSPage<>();
        List<RPTCustomerReceivableSummaryEntity> list = new ArrayList<>();
        RPTCustomer customer = new RPTCustomer();
        int systemId = RptCommonUtils.getSystemId();

        if (new Date().getTime() < queryDate.getTime()) {
            return page;
        }
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        if (search.getCustomerId() != null && search.getCustomerId() != 0) {
            customer.setId(search.getCustomerId());
        }

        if (search.getPaymentType() != null && search.getPaymentType() != 0) {
            customer.getPaymentType().setValue(search.getPaymentType().toString());
        }
        Set<Long> customerIdSet;
        Set<Long> salesIds = new HashSet<>();
        List<RPTCustomer> customerList;
        entityPage.setPageNo(search.getPageNo());
        entityPage.setPageSize(search.getPageSize());
        if (search.getSalesId() != null && search.getSalesId() != 0 ){
            if (search.getCustomerId() != null && search.getCustomerId() > 0) {
                    entityPage = msCustomerService.findCustomerListWithCodeNamePaySaleContract(entityPage, customer);
            }else {
                List<RPTCustomerSalesMappingEntity> customerSalesMappingEntityList = customerReceivableSummaryRptMapper.getCustomerSalesChargeList(systemId, search.getSalesId(),search.getSubFlag());
                List<Long> salesIdList =customerSalesMappingEntityList.stream().map(RPTCustomerSalesMappingEntity::getSalesId).distinct().collect(Collectors.toList());
                entityPage = msCustomerService.findListBySalesIdsForRPT(entityPage, salesIdList);
            }
        }else {
            entityPage = msCustomerService.findCustomerListWithCodeNamePaySaleContract(entityPage, customer);
        }
        customerList = entityPage.getList();
        List<Long> customerIds = Lists.newArrayList();
        if(!CollectionUtils.isEmpty(customerList)) {
            customerIds = customerList.stream().map(RPTCustomer::getId).distinct().collect(Collectors.toList());
        }
        List<RPTCustomerReceivableSummaryEntity> customerReceivableByList = new ArrayList<>();
        List<RPTCustomerReceivableSummaryEntity> preBalanceList = new ArrayList<>();
        List<RPTCustomerReceivableSummaryEntity> currentCreditList = new ArrayList<>();
        if (!CollectionUtils.isEmpty(customerIds)) {
            if (search.getProductCategoryIds().size() == 0) {
                List<Long> productCategoryIds = new ArrayList<>();
                productCategoryIds.add(0L);
                search.setProductCategoryIds(productCategoryIds);
            }
            customerReceivableByList = customerReceivableSummaryRptMapper.getCustomerReceivableByPage(yearMonth, customerIds, search.getProductCategoryIds(), systemId);
            preBalanceList = customerReceivableSummaryRptMapper.getCustomerBalanceData(yearMonth, customerIds, search.getProductCategoryIds(), systemId);
            currentCreditList = customerReceivableSummaryRptMapper.getCurrentCreditList(customerIds);
        }
        Map<Long, RPTCustomerReceivableSummaryEntity> customerReceivableMap = customerReceivableByList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTCustomerReceivableSummaryEntity> preBalanceMap = preBalanceList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTCustomerReceivableSummaryEntity> currentCreditMaps = currentCreditList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));

        RPTDict itemPaymentType;
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");
        RPTCustomerReceivableSummaryEntity entity;
        for (RPTCustomer rptCustomer : customerList) {
            entity = new RPTCustomerReceivableSummaryEntity();
            entity.setCustomerId(rptCustomer.getId());
            entity.setCustomerCode(rptCustomer.getCode());
            entity.setCustomerName(rptCustomer.getName());
            entity.setContractDate(rptCustomer.getContractDate());
            entity.setSalesMan("");
            itemPaymentType = rptCustomer.getPaymentType();
            if (itemPaymentType != null && itemPaymentType.getValue() != null) {
                entity.setPaymentType(paymentTypeMap.get(itemPaymentType.getValue()));
            }

            if (rptCustomer.getSales() != null && rptCustomer.getSales().getId() != null) {
                entity.setSalesId(rptCustomer.getSales().getId());
                salesIds.add(rptCustomer.getSales().getId());
            }
            RPTCustomerReceivableSummaryEntity customerReceivableEntity = customerReceivableMap.get(rptCustomer.getId());
            RPTCustomerReceivableSummaryEntity preBalanceEntity = preBalanceMap.get(rptCustomer.getId());
            RPTCustomerReceivableSummaryEntity currentCreditEntity = currentCreditMaps.get(rptCustomer.getId());
            if (customerReceivableEntity != null) {
                entity.setNewQty(customerReceivableEntity.getNewQty());
                entity.setFinishQty(customerReceivableEntity.getFinishQty());
                entity.setReturnQty(customerReceivableEntity.getReturnQty());
                entity.setCancelQty(customerReceivableEntity.getCancelQty());
                entity.setNoFinishQty(customerReceivableEntity.getNoFinishQty());
                entity.setPreNoFinishQty(customerReceivableEntity.getPreNoFinishQty());

            }
            if (preBalanceEntity != null) {
                entity.setRechargeAmount(preBalanceEntity.getRechargeAmount());
                entity.setDiffCharge(preBalanceEntity.getDiffCharge());
                entity.setBlockAmount(preBalanceEntity.getBlockAmount());
                entity.setPreBalance(preBalanceEntity.getPreBalance());
                entity.setBalance(preBalanceEntity.getBalance());
                entity.setCustomerTimeLinessCharge(preBalanceEntity.getCustomerTimeLinessCharge());
                entity.setOrderPaymentAmount(preBalanceEntity.getOrderPaymentAmount());
                entity.setCustomerUrgentCharge(preBalanceEntity.getCustomerUrgentCharge());
                entity.setPraiseFee(preBalanceEntity.getPraiseFee());

            }

            if (currentCreditEntity != null) {
                entity.setCurrentCredit(currentCreditEntity.getCurrentCredit());
                entity.setCurrentDeposit(currentCreditEntity.getCurrentDeposit());
            }
            list.add(entity);

        }

        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));
        for (RPTCustomerReceivableSummaryEntity rptEntity : list) {
            if (rptEntity.getSalesId() != null) {
                rptEntity.setSalesMan(userNameMap.get(rptEntity.getSalesId()));
            }
        }
        list = list.stream().sorted(Comparator.comparing(RPTCustomerReceivableSummaryEntity::getCustomerName)).collect(Collectors.toList());
        page.setList(list);
        page.setPageSize(entityPage.getPageSize());
        page.setPageNo(entityPage.getPageNo());
        page.setRowCount(entityPage.getRowCount());
        page.setPageCount(entityPage.getPageCount());
        return page;

    }

    public List<RPTCustomerReceivableSummaryEntity> getCustomerReceivableMonth(RPTCustomerOrderPlanDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        List<RPTCustomerReceivableSummaryEntity> list = new ArrayList<>();
        int systemId = RptCommonUtils.getSystemId();
        Set<Long> customerIdSet = Sets.newHashSet();
        Map<Long, RPTCustomer> customerMap = new HashMap<>();
            if (search.getCustomerId() != null && search.getCustomerId() != 0) {
                customerIdSet.add(search.getCustomerId());
                String[] fieldsArray = new String[]{"id", "name", "code", "salesId", "paymentType", "contractDate"};
                customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIdSet), Arrays.asList(fieldsArray));
            }else if(search.getSalesId() != null && search.getSalesId() != 0 ) {
                List<RPTCustomerSalesMappingEntity> customerSalesMappingEntityList = customerReceivableSummaryRptMapper.getCustomerSalesChargeList(systemId, search.getSalesId(), search.getSubFlag());
                customerIdSet = customerSalesMappingEntityList.stream().map(RPTCustomerSalesMappingEntity::getCustomerId).collect(Collectors.toSet());
                String[] fieldsArray = new String[]{"id", "name", "code", "salesId", "paymentType", "contractDate"};
                customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIdSet), Arrays.asList(fieldsArray));
            }else {
                List<RPTCustomer> customerList = msCustomerService.findAllCustomerList();
                for (RPTCustomer item : customerList) {
                    customerIdSet.add(item.getId());
                    customerMap.put(item.getId(), item);
                }
        }

        Set<Long> salesIds = new HashSet<>();
        List<Long> customerIds = Lists.newArrayList(customerIdSet);
        List<RPTCustomerReceivableSummaryEntity> customerReceivableByList = new ArrayList<>();
        List<RPTCustomerReceivableSummaryEntity> preBalanceList = new ArrayList<>();
        List<RPTCustomerReceivableSummaryEntity> currentCreditList = new ArrayList<>();
        if (!customerIds.isEmpty()) {
            if (search.getProductCategoryIds().size() == 0) {
                List<Long> productCategoryIds = new ArrayList<>();
                productCategoryIds.add(0L);
                search.setProductCategoryIds(productCategoryIds);
            }
            customerReceivableByList = customerReceivableSummaryRptMapper.getCustomerReceivableByPage(yearMonth, customerIds.size() <= 100 ? customerIds : null, search.getProductCategoryIds(), systemId);
            preBalanceList = customerReceivableSummaryRptMapper.getCustomerBalanceData(yearMonth, customerIds.size() <= 100 ? customerIds : null, search.getProductCategoryIds(), systemId);
            currentCreditList = customerReceivableSummaryRptMapper.getCurrentCreditList(customerIds.size() <= 100 ? customerIds : null);

        }

        Map<Long, RPTCustomerReceivableSummaryEntity> customerReceivableMap = customerReceivableByList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTCustomerReceivableSummaryEntity> preBalanceMap = preBalanceList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTCustomerReceivableSummaryEntity> currentCreditMaps = currentCreditList.stream().collect(Collectors.toMap(RPTCustomerReceivableSummaryEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));

        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");
        RPTDict itemPaymentType;
        RPTCustomerReceivableSummaryEntity entity;
        for (long customerId : customerIds) {
            RPTCustomer rptCustomer = customerMap.get(customerId);
            if (rptCustomer != null) {
                entity = new RPTCustomerReceivableSummaryEntity();
                entity.setCustomerId(rptCustomer.getId());
                entity.setCustomerCode(rptCustomer.getCode());
                entity.setCustomerName(rptCustomer.getName());
                entity.setContractDate(rptCustomer.getContractDate());
                entity.setSalesMan("");
                itemPaymentType = rptCustomer.getPaymentType();
                if (itemPaymentType != null && itemPaymentType.getValue() != null) {
                    entity.setPaymentType(paymentTypeMap.get(itemPaymentType.getValue()));
                }


                if (rptCustomer.getSales() != null && rptCustomer.getSales().getId() != null) {
                    entity.setSalesId(rptCustomer.getSales().getId());
                    salesIds.add(rptCustomer.getSales().getId());
                }
                RPTCustomerReceivableSummaryEntity customerReceivableEntity = customerReceivableMap.get(rptCustomer.getId());
                RPTCustomerReceivableSummaryEntity preBalanceEntity = preBalanceMap.get(rptCustomer.getId());
                RPTCustomerReceivableSummaryEntity currentCreditEntity = currentCreditMaps.get(rptCustomer.getId());
                if (customerReceivableEntity != null) {
                    entity.setNewQty(customerReceivableEntity.getNewQty());
                    entity.setFinishQty(customerReceivableEntity.getFinishQty());
                    entity.setReturnQty(customerReceivableEntity.getReturnQty());
                    entity.setCancelQty(customerReceivableEntity.getCancelQty());
                    entity.setNoFinishQty(customerReceivableEntity.getNoFinishQty());
                    entity.setPreNoFinishQty(customerReceivableEntity.getPreNoFinishQty());

                }
                if (preBalanceEntity != null) {
                    entity.setRechargeAmount(preBalanceEntity.getRechargeAmount());
                    entity.setDiffCharge(preBalanceEntity.getDiffCharge());
                    entity.setBlockAmount(preBalanceEntity.getBlockAmount());
                    entity.setPreBalance(preBalanceEntity.getPreBalance());
                    entity.setBalance(preBalanceEntity.getBalance());
                    entity.setCustomerTimeLinessCharge(preBalanceEntity.getCustomerTimeLinessCharge());
                    entity.setOrderPaymentAmount(preBalanceEntity.getOrderPaymentAmount());
                    entity.setCustomerUrgentCharge(preBalanceEntity.getCustomerUrgentCharge());
                    entity.setPraiseFee(preBalanceEntity.getPraiseFee());

                }

                if (currentCreditEntity != null) {
                    entity.setCurrentCredit(currentCreditEntity.getCurrentCredit());
                    entity.setCurrentDeposit(currentCreditEntity.getCurrentDeposit());
                }


                list.add(entity);
            }
        }

        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));
        for (RPTCustomerReceivableSummaryEntity rptEntity : list) {
            if (rptEntity.getSalesId() != null) {
                rptEntity.setSalesMan(userNameMap.get(rptEntity.getSalesId()));
            }
        }

        list = list.stream().sorted(Comparator.comparing(RPTCustomerReceivableSummaryEntity::getCustomerName)).collect(Collectors.toList());

        return list;

    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        if (searchCondition.getStartDate() != null) {
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getStartDate()));
            Integer selectedMonth = DateUtils.getMonth(new Date(searchCondition.getStartDate()));
            Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
            Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
            int systemId = RptCommonUtils.getSystemId();
            Integer rowCountA = 0;
            if (new Date().getTime() < queryDate.getTime()) {
                return false;
            }
            if (searchCondition.getCustomerId() != null) {
               rowCountA = customerReceivableSummaryRptMapper.hasReportList(searchCondition.getCustomerId());
            }

            Integer rowCountB =  customerReceivableSummaryRptMapper.hasReportData(yearMonth,searchCondition.getCustomerId(),searchCondition.getProductCategoryIds(),systemId);
            result = rowCountA + rowCountB > 0;

        }
        return result;
    }

    public SXSSFWorkbook customerReceivableRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTCustomerReceivableSummaryEntity> list = getCustomerReceivableMonth(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            boolean result = true;
            if (searchCondition.getProductCategoryIds().size() > 0 && searchCondition.getProductCategoryIds().get(0) != 0L) {
                result = false;
            }
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            if (result) {
                Row titleRow = xSheet.createRow(rowIndex++);
                titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
                ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
                xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 22));

                Row headFirstRow = xSheet.createRow(rowIndex++);
                headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
                ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
                ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
                ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
                ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
                ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5));
                ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

                xSheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 11));
                ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "订单情况");

                xSheet.addMergedRegion(new CellRangeAddress(1, 1, 12, 20));
                ExportExcel.createCell(headFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "消费情况");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 21, 21));
                ExportExcel.createCell(headFirstRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "信用额度");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 22, 22));
                ExportExcel.createCell(headFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户押金");

                Row headSecondRow = xSheet.createRow(rowIndex++);
                headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

                ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月未完成单");
                ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月接单");
                ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单");
                ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退单");
                ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月取消单");
                ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月未完成单");

                ExportExcel.createCell(headSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月消费余额");
                ExportExcel.createCell(headSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月已收款");
                ExportExcel.createCell(headSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月应收款");
                ExportExcel.createCell(headSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "对账差异单");
                ExportExcel.createCell(headSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月时效费");
                ExportExcel.createCell(headSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月加急费");
                ExportExcel.createCell(headSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月好评费");
                ExportExcel.createCell(headSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月消费余额");
                ExportExcel.createCell(headSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月冻结金额");

                xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

                // 写入数据
                Cell dataCell = null;
                if (list != null && list.size() > 0) {

                    int preNoFinishQty = 0;
                    int newQty = 0;
                    int finishQty = 0;
                    int returnQty = 0;
                    int cancelQty = 0;
                    int noFinishQty = 0;

                    double preBalance = 0.00;
                    double rechargeAmount = 0.00;
                    double orderPaymentAmount = 0.00;
                    double diffCharge = 0.00;
                    double customerTimeLinessCharge = 0.00;
                    double customerUrgentCharge = 0.00;
                    double balance = 0.00;
                    double blockAmount = 0.00;
                    double praiseFee = 0.00;

                    double currentCredit = 0.00;
                    double currentDeposit = 0.00;

                    int rowsCount = list.size();
                    for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                        RPTCustomerReceivableSummaryEntity rowData = list.get(dataRowIndex);

                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 0;

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getContractDate(), "yyyyMM"));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerCode());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSalesMan());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPreNoFinishQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getNewQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getFinishQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReturnQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCancelQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getNoFinishQty());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPreBalance());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRechargeAmount());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderPaymentAmount());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getDiffCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerTimeLinessCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerUrgentCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPraiseFee());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBalance());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBlockAmount());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCurrentCredit());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCurrentDeposit());

                        preNoFinishQty = preNoFinishQty + rowData.getPreNoFinishQty();
                        newQty = newQty + rowData.getNewQty();
                        finishQty = finishQty + rowData.getFinishQty();
                        returnQty = returnQty + rowData.getReturnQty();
                        cancelQty = cancelQty + rowData.getCancelQty();
                        noFinishQty = noFinishQty + rowData.getNoFinishQty();

                        preBalance = preBalance + rowData.getPreBalance();
                        rechargeAmount = rechargeAmount + rowData.getRechargeAmount();
                        orderPaymentAmount = orderPaymentAmount + rowData.getOrderPaymentAmount();
                        diffCharge = diffCharge + rowData.getDiffCharge();
                        customerTimeLinessCharge = customerTimeLinessCharge + rowData.getCustomerTimeLinessCharge();
                        customerUrgentCharge = customerUrgentCharge + rowData.getCustomerUrgentCharge();
                        praiseFee = praiseFee + rowData.getPraiseFee();
                        balance = balance + rowData.getBalance();
                        blockAmount = blockAmount + rowData.getBlockAmount();

                        currentCredit = currentCredit + rowData.getCurrentCredit();
                        currentDeposit = currentDeposit + rowData.getCurrentDeposit();
                    }

                    Row sumRow = xSheet.createRow(rowIndex++);
                    sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 5));
                    ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                    ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, preNoFinishQty);
                    ExportExcel.createCell(sumRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, newQty);
                    ExportExcel.createCell(sumRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, finishQty);
                    ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, returnQty);
                    ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cancelQty);
                    ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, noFinishQty);

                    ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, preBalance);
                    ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rechargeAmount);
                    ExportExcel.createCell(sumRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderPaymentAmount);
                    ExportExcel.createCell(sumRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, diffCharge);
                    ExportExcel.createCell(sumRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimeLinessCharge);
                    ExportExcel.createCell(sumRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerUrgentCharge);
                    ExportExcel.createCell(sumRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
                    ExportExcel.createCell(sumRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, balance);
                    ExportExcel.createCell(sumRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, blockAmount);

                    ExportExcel.createCell(sumRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, currentCredit);
                    ExportExcel.createCell(sumRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, currentDeposit);
                }
            } else {
                Row titleRow = xSheet.createRow(rowIndex++);
                titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
                ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
                xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 15));

                Row headFirstRow = xSheet.createRow(rowIndex++);
                headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
                ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
                ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
                ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
                ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
                ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");

                xSheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5));
                ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");

                xSheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 11));
                ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "订单情况");

                xSheet.addMergedRegion(new CellRangeAddress(1, 1, 12, 16));
                ExportExcel.createCell(headFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "消费情况");



                Row headSecondRow = xSheet.createRow(rowIndex++);
                headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

                ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月未完成单");
                ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月接单");
                ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单");
                ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退单");
                ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月取消单");
                ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月未完成单");


                ExportExcel.createCell(headSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月应收款");
                ExportExcel.createCell(headSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "对账差异单");
                ExportExcel.createCell(headSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月时效费");
                ExportExcel.createCell(headSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月加急费");
                ExportExcel.createCell(headSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月好评费");


                xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

                // 写入数据
                Cell dataCell = null;
                if (list != null && list.size() > 0) {

                    int preNoFinishQty = 0;
                    int newQty = 0;
                    int finishQty = 0;
                    int returnQty = 0;
                    int cancelQty = 0;
                    int noFinishQty = 0;

                    double orderPaymentAmount = 0.00;
                    double diffCharge = 0.00;
                    double customerTimeLinessCharge = 0.00;
                    double customerUrgentCharge = 0.00;
                    double praiseFee = 0.00;


                    int rowsCount = list.size();
                    for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                        RPTCustomerReceivableSummaryEntity rowData = list.get(dataRowIndex);

                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 0;

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel()));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getContractDate(), "yyyyMM"));
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerCode());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSalesMan());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPreNoFinishQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getNewQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getFinishQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReturnQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCancelQty());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getNoFinishQty());


                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderPaymentAmount());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getDiffCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerTimeLinessCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerUrgentCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPraiseFee());


                        preNoFinishQty = preNoFinishQty + rowData.getPreNoFinishQty();
                        newQty = newQty + rowData.getNewQty();
                        finishQty = finishQty + rowData.getFinishQty();
                        returnQty = returnQty + rowData.getReturnQty();
                        cancelQty = cancelQty + rowData.getCancelQty();
                        noFinishQty = noFinishQty + rowData.getNoFinishQty();

                        orderPaymentAmount = orderPaymentAmount + rowData.getOrderPaymentAmount();
                        diffCharge = diffCharge + rowData.getDiffCharge();
                        customerTimeLinessCharge = customerTimeLinessCharge + rowData.getCustomerTimeLinessCharge();
                        customerUrgentCharge = customerUrgentCharge + rowData.getCustomerUrgentCharge();
                        praiseFee = praiseFee + rowData.getPraiseFee();

                    }

                    Row sumRow = xSheet.createRow(rowIndex++);
                    sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 5));
                    ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                    ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, preNoFinishQty);
                    ExportExcel.createCell(sumRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, newQty);
                    ExportExcel.createCell(sumRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, finishQty);
                    ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, returnQty);
                    ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cancelQty);
                    ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, noFinishQty);


                    ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderPaymentAmount);
                    ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, diffCharge);
                    ExportExcel.createCell(sumRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimeLinessCharge);
                    ExportExcel.createCell(sumRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerUrgentCharge);
                    ExportExcel.createCell(sumRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
                }
            }

        } catch (Exception e) {
            log.error("【getCustomerReceivableSummaryMonth.customerReceivableRptExport】客户订单汇总写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
