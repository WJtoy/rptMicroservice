package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.GlobalMappingSalesSubFlagEnum;
import com.kkl.kklplus.entity.rpt.*;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.entity.CustomerPerformanceRptEntity;
import com.kkl.kklplus.provider.rpt.entity.GradeQtyEntity;
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;

import com.kkl.kklplus.provider.rpt.mapper.CustomerPerformanceRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.catalina.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.codehaus.groovy.ast.GenericsType;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.lang.reflect.Method;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerPerformanceRptService extends RptBaseService {

    @Autowired
    private CustomerPerformanceRptMapper customerPerformanceRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;


    public List<RPTSalesPerfomanceEntity> getSalesPerformanceByList(RPTCustomerOrderPlanDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        List<RPTSalesPerfomanceEntity> finishQtyList = customerPerformanceRptMapper.getCustomerPerformanceByList(yearMonth,search.getProductCategoryIds(),systemId,search.getSalesId(),search.getSubFlag());
        List<RPTSalesPerfomanceEntity> noFinishQtyList = customerPerformanceRptMapper.getSalesNoFinishOrderQty(yearMonth,search.getProductCategoryIds(),systemId,search.getSalesId(),search.getSubFlag());


        Set<Long> salesIdSet = Sets.newHashSet();
        if(null != search.getSalesId() && search.getSalesId() != 0L){
            salesIdSet.add(search.getSalesId());
            if (search.getSubFlag()!=null && search.getSubFlag()!=0 && search.getSubFlag()== GlobalMappingSalesSubFlagEnum.MANAGER.getValue()){
                List<RPTCustomerSalesMappingEntity> salesCustomerList = customerPerformanceRptMapper.getCustomerSalesChargeList(systemId,search.getSalesId());
                for (RPTCustomerSalesMappingEntity salesEntity : salesCustomerList) {
                    if(salesEntity.getSalesId() != 0L){
                        salesIdSet.add(salesEntity.getSalesId());
                    }
                }
            }
        }else {
            List<RPTCustomerSalesMappingEntity> salesCustomerList = customerPerformanceRptMapper.getCustomerSalesList(systemId);
            for (RPTCustomerSalesMappingEntity salesEntity : salesCustomerList) {
                if(salesEntity.getSalesId() != 0L){
                    salesIdSet.add(salesEntity.getSalesId());
                }
            }
        }
        Map<Long, RPTSalesPerfomanceEntity> salesPerformanceEntityMap = finishQtyList.stream().collect(Collectors.toMap(RPTSalesPerfomanceEntity::getSalesId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTSalesPerfomanceEntity> noFinishQtyMap = noFinishQtyList.stream().collect(Collectors.toMap(RPTSalesPerfomanceEntity::getSalesId, Function.identity(), (key1, key2) -> key2));

        List<Long> salesIds = Lists.newArrayList(salesIdSet);
        Map<Long, String> salesMap = MSUserUtils.getNamesByUserIds(salesIds);
        RPTSalesPerfomanceEntity entity;
        RPTSalesPerfomanceEntity rptEntity;
        RPTSalesPerfomanceEntity noFinishEntity;

        List<RPTSalesPerfomanceEntity> list = Lists.newArrayList();
        for(long salesId : salesIds){
            rptEntity = new RPTSalesPerfomanceEntity();
            String name = salesMap.get(salesId);
            entity = salesPerformanceEntityMap.get(salesId);
            noFinishEntity =  noFinishQtyMap.get(salesId);
            if(entity != null){
                rptEntity.setCreateQty(entity.getCreateQty());
                rptEntity.setFinishQty(entity.getFinishQty());
                rptEntity.setReturnQty(entity.getReturnQty());
                rptEntity.setCancelQty(entity.getCancelQty());
            }

            if(noFinishEntity != null){
                rptEntity.setNoFinishQty(noFinishEntity.getNoFinishQty());
            }
            rptEntity.setSalesName(name);
            rptEntity.setSalesId(salesId);
            list.add(rptEntity);

        }

        list = list.stream().sorted(Comparator.comparing(RPTSalesPerfomanceEntity::getCreateQty).reversed()).collect(Collectors.toList());

        return list;
    }

    public List<RPTSalesPerfomanceEntity> getCustomerPerformanceList(RPTCustomerOrderPlanDailySearch search) {
        Long salesId = null;
        Integer subFlag = null;
        if(search.getSalesId() != null ){
            salesId = search.getSalesId();
        }
        if (search.getSubFlag() != null){
            subFlag =search.getSubFlag();

        }
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        List<RPTSalesPerfomanceEntity> finishQtyList = customerPerformanceRptMapper.getSalesManAchievementDataNew(yearMonth,search.getProductCategoryIds(),systemId,salesId,subFlag);
        List<RPTSalesPerfomanceEntity> noFinishQtyList = customerPerformanceRptMapper.getCustomerNoFinishOrderQtyNew(yearMonth, salesId,search.getProductCategoryIds(),systemId,subFlag);
        List<RPTCustomerSalesMappingEntity> salesCustomerList = customerPerformanceRptMapper.getCustomerSalesList(systemId);


        Set<Long> customerIdSet = Sets.newHashSet();
        for (RPTCustomerSalesMappingEntity customerEntity : salesCustomerList) {
            customerIdSet.add(customerEntity.getCustomerId());
        }


        Map<Long, RPTSalesPerfomanceEntity> customerPerformanceEntityMap = finishQtyList.stream().collect(Collectors.toMap(RPTSalesPerfomanceEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));
        Map<Long, RPTSalesPerfomanceEntity> noFinishQtyMap = noFinishQtyList.stream().collect(Collectors.toMap(RPTSalesPerfomanceEntity::getCustomerId, Function.identity(), (key1, key2) -> key2));

        List<Long> customerIds = Lists.newArrayList(customerIdSet);
        String[] fieldsArray = new String[]{"id", "name","code","salesId"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));
        RPTSalesPerfomanceEntity entity;
        RPTSalesPerfomanceEntity rptEntity;
        RPTSalesPerfomanceEntity noFinishEntity;
        String userName;
        List<RPTSalesPerfomanceEntity> list = Lists.newArrayList();

        for(long customerId : customerIds){
            RPTCustomer customer = customerMap.get(customerId);
            if(salesId != null) {
                if (customer != null && customer.getSales() != null && customer.getSales().getId() != null && salesId.equals(customer.getSales().getId())) {
                    rptEntity = new RPTSalesPerfomanceEntity();
                    rptEntity.setCustomerId(customerId);
                    rptEntity.setCustomerCode(customer.getCode());
                    rptEntity.setCustomerName(customer.getName());
                    userName = userNameMap.get(customer.getSales().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        rptEntity.setSalesName(userName);
                    }
                    rptEntity.setSalesId(customer.getSales().getId());
                    entity = customerPerformanceEntityMap.get(customerId);
                    noFinishEntity = noFinishQtyMap.get(customerId);
                    if (entity != null) {
                        rptEntity.setCreateQty(entity.getCreateQty());
                        rptEntity.setFinishQty(entity.getFinishQty());
                        rptEntity.setReturnQty(entity.getReturnQty());
                        rptEntity.setCancelQty(entity.getCancelQty());
                    }
                    if(noFinishEntity != null){
                        rptEntity.setNoFinishQty(noFinishEntity.getNoFinishQty());
                    }
                    list.add(rptEntity);
                }
            }else {
                if (customer != null && customer.getSales() != null && customer.getSales().getId() != null ) {
                    rptEntity = new RPTSalesPerfomanceEntity();
                    rptEntity.setCustomerId(customerId);
                    rptEntity.setCustomerCode(customer.getCode());
                    rptEntity.setCustomerName(customer.getName());
                    rptEntity.setSalesId(customer.getSales().getId());
                    userName = userNameMap.get(customer.getSales().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        rptEntity.setSalesName(userName);
                    }
                    entity = customerPerformanceEntityMap.get(customerId);
                    noFinishEntity = noFinishQtyMap.get(customerId);
                    if (entity != null) {
                        rptEntity.setCreateQty(entity.getCreateQty());
                        rptEntity.setFinishQty(entity.getFinishQty());
                        rptEntity.setReturnQty(entity.getReturnQty());
                        rptEntity.setCancelQty(entity.getCancelQty());
                        rptEntity.setNoFinishQty(entity.getCreateQty() - entity.getFinishQty() - entity.getCancelQty() - entity.getReturnQty());
                    }

                    if(noFinishEntity != null){
                        rptEntity.setNoFinishQty(noFinishEntity.getNoFinishQty());
                    }

                    list.add(rptEntity);
                }
            }
        }
        list = list.stream().sorted(Comparator.comparing(RPTSalesPerfomanceEntity::getCreateQty).reversed()).collect(Collectors.toList());
        return list;
    }

    public List<RPTSalesPerfomanceEntity> getSalesPerformanceList(RPTCustomerOrderPlanDailySearch search){
        List<RPTSalesPerfomanceEntity>  list = getSalesPerformanceByList(search);
        List<RPTSalesPerfomanceEntity>  customerPerformanceList = getCustomerPerformanceList(search);
        Map<Long, List<RPTSalesPerfomanceEntity>> salesCustomerEntityMap = customerPerformanceList.stream().collect(Collectors.groupingBy(RPTSalesPerfomanceEntity::getSalesId));

        for(RPTSalesPerfomanceEntity entity: list){
            List<RPTSalesPerfomanceEntity>  customerList  =  salesCustomerEntityMap.get(entity.getSalesId());
            if(customerList != null) {
                customerList = customerList.stream().sorted(Comparator.comparing(RPTSalesPerfomanceEntity::getCreateQty).reversed()).collect(Collectors.toList());
            }
            entity.setItemList(customerList);
        }
        list = list.stream().sorted(Comparator.comparing(RPTSalesPerfomanceEntity::getCreateQty).reversed()).collect(Collectors.toList());

        return list;
    }




    public List<CustomerPerformanceRptEntity> insertSalesPerformance(int selectedYear, int selectedMonth) {
        List<CustomerPerformanceRptEntity> result = Lists.newArrayList();
        Date queryDate =  DateUtils.getStartOfDay(DateUtils.getDate(selectedYear, selectedMonth, 1));
        Integer yearMonth = generateYearMonth(queryDate);
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        int currentYearMonth = generateYearMonth(new Date());
        Date endDate;
        if (yearMonth == currentYearMonth) {
            endDate = DateUtils.getStartOfDay(new Date());
        } else {
            endDate = DateUtils.addMonth(queryDate, 1);
        }
        int systemId = RptCommonUtils.getSystemId();
        long beginCreatDate = queryDate.getTime();
        long endCreatDate = endDate.getTime();
        Map<String, CustomerPerformanceRptEntity> finishQtyMap = new HashMap<>();
        List<CustomerPerformanceRptEntity> finishQtyList = customerPerformanceRptMapper.getFinishQtyList(beginCreatDate, endCreatDate, systemId,quarter);

        List<CustomerPerformanceRptEntity> returnQtyList = customerPerformanceRptMapper.getReturnQtyList(beginCreatDate, endCreatDate, systemId,quarter);
        List<CustomerPerformanceRptEntity> cancelQtyList = customerPerformanceRptMapper.getCancelQtyList(beginCreatDate, endCreatDate, systemId,quarter);
        List<CustomerPerformanceRptEntity> createQtyList = customerPerformanceRptMapper.getCreateQtyList(beginCreatDate, endCreatDate, systemId,quarter);


        for (CustomerPerformanceRptEntity entity : createQtyList) {
            String createKey = entity.getCustomerId() + "," + entity.getProductCategoryId();
            finishQtyMap.put(createKey, entity);

        }
        for (CustomerPerformanceRptEntity entity : finishQtyList) {
            String finishKey = entity.getCustomerId() + "," + entity.getProductCategoryId();
            if (finishQtyMap.containsKey(finishKey)) {
                finishQtyMap.get(finishKey).setFinishQty(entity.getFinishQty());
            } else {
                finishQtyMap.put(finishKey, entity);
            }
        }

        for (CustomerPerformanceRptEntity entity : returnQtyList) {
            String returnKey = entity.getCustomerId() + "," + entity.getProductCategoryId();
            if (finishQtyMap.containsKey(returnKey)) {
                finishQtyMap.get(returnKey).setReturnQty(entity.getReturnQty());
            } else {
                finishQtyMap.put(returnKey, entity);
            }
        }
        for (CustomerPerformanceRptEntity entity : cancelQtyList) {
            String cancelKey = entity.getCustomerId() + "," + entity.getProductCategoryId();
            if (finishQtyMap.containsKey(cancelKey)) {
                finishQtyMap.get(cancelKey).setCancelQty(entity.getCancelQty());
            } else {
                finishQtyMap.put(cancelKey, entity);
            }
        }
        for(CustomerPerformanceRptEntity entity : finishQtyMap.values()){
            entity.setYearMonth(yearMonth);
            result.add(entity);
        }
        return result;
    }

    private Map<String, Long> getCustomerFinanceMonthlyIdMap(int systemId, int yearmonth) {
        List<LongThreeTuple> tuples = customerPerformanceRptMapper.getCustomerOrderQtyMonthlyIds(systemId, yearmonth);
        Map<String, Long> tuplesMap = Maps.newHashMap();
        if(tuples != null && !tuples.isEmpty()) {
            for(LongThreeTuple item : tuples ){
                String key = StringUtils.join(item.getBElement(), "%", item.getCElement());
                tuplesMap.put(key,item.getAElement());
            }
            return tuplesMap;
        } else {
            return tuplesMap;
        }
    }


    /**
     * 写入数据
     */
    public void writeYesterdayQty(int selectedYear, int selectedMonth) {
        updateCustomerPerformanceToRptDB(selectedYear, selectedMonth);

    }

    private void updateCustomerPerformanceToRptDB(int selectedYear, int selectedMonth) {
        List<CustomerPerformanceRptEntity> list = insertSalesPerformance(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            Map<String, Long> idMap = getCustomerFinanceMonthlyIdMap(systemId, yearmonth);
            Long primaryKeyId;
            String key;
            for (CustomerPerformanceRptEntity item : list) {
                key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
                primaryKeyId = idMap.get(key);
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    customerPerformanceRptMapper.updateCustomerOrderQtyMonthly(item);
                } else {
                    item.setSystemId(systemId);
                    customerPerformanceRptMapper.insertCustomerPerformance(item);
                }
            }
        }
    }


    public void deleteCustomerPerformanceRptDB(int selectedYear, int selectedMonth){
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        customerPerformanceRptMapper.delete(yearMonth,systemId);
    }

    public void saveCustomerPerformanceRptDB(int selectedYear, int selectedMonth){
        List<CustomerPerformanceRptEntity> list = insertSalesPerformance(selectedYear, selectedMonth);
        int systemId = RptCommonUtils.getSystemId();
        if (!list.isEmpty()) {
            for (CustomerPerformanceRptEntity item : list) {
                item.setSystemId(systemId);
                customerPerformanceRptMapper.insertCustomerPerformance(item);
            }
        }

    }

    private int generateYearMonth(Date date) {
        int selectedYear = DateUtils.getYear(date);
        int selectedMonth = DateUtils.getMonth(date);
        return generateYearMonth(selectedYear, selectedMonth);
    }

    private int generateYearMonth(int selectedYear, int selectedMonth) {
        return StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
    }

    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                int selectedYear;
                int selectedMonth;
                while (beginDate.getTime() <= endDate.getTime()) {
                    selectedYear = DateUtils.getYear(beginDate);
                    selectedMonth = DateUtils.getMonth(beginDate);
                    switch (operationType) {
                        case INSERT:
                            saveCustomerPerformanceRptDB(selectedYear, selectedMonth);
                            break;
                        case INSERT_MISSED_DATA:
                            updateCustomerPerformanceToRptDB(selectedYear, selectedMonth);
                            break;
                        case UPDATE:
                            deleteCustomerPerformanceRptDB(selectedYear, selectedMonth);
                            saveCustomerPerformanceRptDB(selectedYear, selectedMonth);
                            break;
                        case DELETE:
                            deleteCustomerPerformanceRptDB(selectedYear, selectedMonth);
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }

                result = true;
            } catch (Exception e) {
                log.error("CustomerPerformanceRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;

    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        Long salesId = null;
        Integer subFlag = null;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition.getStartDate() != null) {
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getStartDate()));
            Integer selectedMonth = DateUtils.getMonth(new Date(searchCondition.getStartDate()));
            Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
            if (new Date().getTime() < queryDate.getTime()) {
                return false;
            }
            Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
            if(searchCondition.getSalesId() != null ){
                salesId = searchCondition.getSalesId();
            }
            if (searchCondition.getSubFlag()!=null){
                subFlag = searchCondition.getSubFlag();
            }
            Integer rowCount = customerPerformanceRptMapper.hasReportData(yearMonth,searchCondition.getProductCategoryIds(),systemId,salesId,subFlag);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook salesPerformanceMonthRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTSalesPerfomanceEntity> list =  getSalesPerformanceList(searchCondition);
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
            Row headerRow = xSheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "排名");
            ExportExcel.createCell(headerRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");
            ExportExcel.createCell(headerRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");
            ExportExcel.createCell(headerRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单数量");
            ExportExcel.createCell(headerRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工单数量");
            ExportExcel.createCell(headerRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工单数量");
            ExportExcel.createCell(headerRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单数量");
            ExportExcel.createCell(headerRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消单数量");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            Row dataRow = null;
            if (list != null && list.size() > 0) {
                int totalNewQty = 0;
                int totaFinishQty = 0;
                int totalProcessQty = 0;
                int totalReturnQty = 0;
                int totalCancelQty = 0;

                for (int i = 0; i < list.size(); i++) {
                    RPTSalesPerfomanceEntity item = list.get(i);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    //业务员行
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getSalesName());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCreateQty());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getFinishQty());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getNoFinishQty());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getReturnQty());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCancelQty());

                    totalNewQty = totalNewQty + item.getCreateQty();
                    totaFinishQty = totaFinishQty + item.getFinishQty();
                    totalProcessQty = totalProcessQty + item.getNoFinishQty();
                    totalReturnQty = totalReturnQty + item.getReturnQty();
                    totalCancelQty = totalCancelQty + item.getCancelQty();

                    if (item.getItemList() != null) {
                        for (int subIndex = 0; subIndex < item.getItemList().size(); subIndex++) {
                            RPTSalesPerfomanceEntity subItem = item.getItemList().get(subIndex);
                            if (subIndex == 0) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                int rowSpan = item.getItemList().size() - 1;

                                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                if (rowSpan > 0) {
                                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                                }
                                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getSalesName());
                                if (rowSpan > 0) {
                                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                                }

                                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCustomerCode());
                                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCustomerName());
                                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCreateQty());
                                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getFinishQty());
                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getNoFinishQty());
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getReturnQty());
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCancelQty());
                            } else {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCustomerCode());
                                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCustomerName());
                                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCreateQty());
                                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getFinishQty());
                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getNoFinishQty());
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getReturnQty());
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, subItem.getCancelQty());
                            }
                        }

                    }
                }
                dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 3));

                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalNewQty);
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totaFinishQty);
                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalProcessQty);
                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalReturnQty);
                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCancelQty);
            }

        } catch (Exception e) {
            log.error("【CustomerPerformanceRptService.salesPerformanceMonthRptExport】业务业绩报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


    public SXSSFWorkbook customerPerformanceMonthRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTSalesPerfomanceEntity> list = getCustomerPerformanceList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 7));

            //====================================================绘制表头============================================================
            Row headerRow = xSheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");
            ExportExcel.createCell(headerRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");
            ExportExcel.createCell(headerRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单数量");
            ExportExcel.createCell(headerRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工单数量");
            ExportExcel.createCell(headerRow,5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工单数量");
            ExportExcel.createCell(headerRow,6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单数量");
            ExportExcel.createCell(headerRow,7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "取消单数量");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            Row dataRow = null;
            if (list != null && list.size() > 0){
                int totalNewQty = 0;
                int totaFinishQty = 0;
                int totalProcessQty = 0;
                int totalReturnQty = 0;
                int totalCancelQty = 0;
                for (int i=0; i < list.size(); i++) {
                    RPTSalesPerfomanceEntity item = list.get(i);
                    int rowSpan = list.size() - 1;
                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    //业务员行


                    if ( i == 0) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getSalesName());
                        if(rowSpan>0) {
                            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                        }
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCustomerCode());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCustomerName());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCreateQty());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getFinishQty());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getNoFinishQty());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getReturnQty());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCancelQty());

                    totalNewQty = totalNewQty + item.getCreateQty();
                    totaFinishQty = totaFinishQty + item.getFinishQty();
                    totalProcessQty = totalProcessQty + item.getNoFinishQty();
                    totalReturnQty = totalReturnQty + item.getReturnQty();
                    totalCancelQty = totalCancelQty + item.getCancelQty();
                }
                dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 2));

                ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalNewQty);
                ExportExcel.createCell(dataRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totaFinishQty);
                ExportExcel.createCell(dataRow,5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalProcessQty);
                ExportExcel.createCell(dataRow,6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalReturnQty);
                ExportExcel.createCell(dataRow,7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCancelQty);
            }
        } catch (Exception e) {
            log.error("【CustomerPerformanceRptService.customerPerformanceMonthRptExport】业务业绩明细报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}
