package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerReminderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.mapper.CustomerReminderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
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
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerReminderRptService extends RptBaseService{

    @Resource
    private CustomerReminderRptMapper customerReminderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    //获取报表显示数据
    public List<RPTCustomerReminderEntity> getCustomerReminderList(RPTCustomerReminderSearch search) {
        List<RPTCustomerReminderEntity> list = new ArrayList<>();
        if (search != null && search.getCustomerId() != null && search.getCustomerId() > 0
                && search.getSelectedYear() != null && search.getSelectedYear() > 0
                && search.getSelectedMonth() != null && search.getSelectedMonth() > 0) {
            Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
            Date beginDate = DateUtils.getStartOfDay(queryDate);
            String quarter = QuarterUtils.getSeasonQuarter(queryDate);
            Date endDate = DateUtils.addMonth(beginDate, 1);
            Integer days = DateUtils.getDaysOfMonth(queryDate);
            Long customerId = search.getCustomerId();
            List<Long> productCategoryIds = search.getProductCategoryIds();
            long beginDateTime = beginDate.getTime();
            long endDateTime = endDate.getTime();
            int systemId = RptCommonUtils.getSystemId();
            List<RPTCustomerReminderEntity> returnList = customerReminderRptMapper.getCustomerReminderList(customerId, systemId, productCategoryIds, beginDateTime, endDateTime,quarter);
            if(returnList.isEmpty()){
                return list;
            }
            Map<Long, List<RPTCustomerReminderEntity>> groupBy = returnList.stream().collect(Collectors.groupingBy(RPTCustomerReminderEntity::getStatisticsDt));
            Integer orderNewQty, reminderQty, reminderFirstQty, reminderMultipleQty,
                    reminderOrderQty, exceed48hourReminderQty, complete24hourQty, over48ReminderCompletedQty;
            String reminderRate, reminderFirstRate, reminderMultipleRate, exceed48hourReminderRate, complete24hourRate, over48ReminderCompletedRate;
            NumberFormat numberFormat = NumberFormat.getInstance();
            numberFormat.setMaximumFractionDigits(2);
            long key;
            String customerName = "";

            RPTCustomer customer = msCustomerService.get(customerId);

            if(customer != null && customer.getName() != null){
                customerName = customer.getName();
            }
            for (int i = 0; i < days; i++) {
                RPTCustomerReminderEntity rpt = new RPTCustomerReminderEntity();
                key = DateUtils.addDays(beginDate,i).getTime();
                if(groupBy.get(key)== null){
                    rpt.setStatisticsDate(new Date(key));
                    rpt.setCustomerName(customerName);
                }else {
                    rpt.setStatisticsDate(new Date(key));
                    orderNewQty = 0; reminderQty = 0; reminderFirstQty = 0; reminderMultipleQty = 0;
                    reminderOrderQty = 0; exceed48hourReminderQty = 0; complete24hourQty = 0; over48ReminderCompletedQty = 0;
                    for(RPTCustomerReminderEntity entity : groupBy.get(key)){
                        orderNewQty += entity.getOrderNewQty();
                        reminderQty += entity.getReminderQty();
                        reminderFirstQty += entity.getReminderFirstQty();
                        reminderMultipleQty += entity.getReminderMultipleQty();
                        reminderOrderQty += entity.getReminderOrderQty();
                        exceed48hourReminderQty += entity.getExceed48hourReminderQty();
                        complete24hourQty += entity.getComplete24hourQty();
                        over48ReminderCompletedQty += entity.getOver48ReminderCompletedQty();
                    }
                    rpt.setOrderNewQty(orderNewQty);
                    rpt.setReminderQty(reminderQty);
                    rpt.setReminderFirstQty(reminderFirstQty);
                    rpt.setReminderMultipleQty(reminderMultipleQty);
                    rpt.setReminderOrderQty(reminderOrderQty);
                    rpt.setExceed48hourReminderQty(exceed48hourReminderQty);
                    rpt.setComplete24hourQty(complete24hourQty);
                    rpt.setOver48ReminderCompletedQty(over48ReminderCompletedQty);

                    orderNewQty = rpt.getOrderNewQty();
                    reminderQty = rpt.getReminderQty();

                    rpt.setCustomerName(customerName);

                    if (orderNewQty != null && orderNewQty != 0) {
                        reminderRate = numberFormat.format((float) reminderQty / orderNewQty * 100);
                        reminderFirstQty = rpt.getReminderFirstQty();
                        reminderFirstRate = numberFormat.format((float) reminderFirstQty / orderNewQty * 100);
                        reminderMultipleQty = rpt.getReminderMultipleQty();
                        reminderMultipleRate = numberFormat.format((float) reminderMultipleQty / orderNewQty * 100);

                        exceed48hourReminderQty = rpt.getExceed48hourReminderQty();
                        exceed48hourReminderRate = numberFormat.format((float) exceed48hourReminderQty / orderNewQty * 100);
                        rpt.setReminderRate(reminderRate);
                        rpt.setReminderFirstRate(reminderFirstRate);
                        rpt.setReminderMultipleRate(reminderMultipleRate);
                        rpt.setExceed48hourReminderRate(exceed48hourReminderRate);
                    }
                    reminderOrderQty = rpt.getReminderOrderQty();
                    if (reminderOrderQty != null && reminderOrderQty != 0) {
                        complete24hourQty = rpt.getComplete24hourQty();
                        complete24hourRate = numberFormat.format((float) complete24hourQty / reminderOrderQty * 100);
                        over48ReminderCompletedQty = rpt.getOver48ReminderCompletedQty();
                        over48ReminderCompletedRate = numberFormat.format((float) over48ReminderCompletedQty / reminderOrderQty * 100);

                        rpt.setComplete24hourRate(complete24hourRate);
                        rpt.setOver48ReminderCompletedRate(over48ReminderCompletedRate);
                    }
                }
                list.add(rpt);
            }

        }

        return list;
    }

    //获取写入中间表数据
    private List<RPTCustomerReminderEntity> getCustomerReminderListFromWebDB(Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(beginDate);
        Long beginDt = beginDate.getTime();
        Long endDt = endDate.getTime();
        List<RPTCustomerReminderEntity> list = new ArrayList<>();

        List<RPTCustomerReminderEntity> orderNewQtyList = customerReminderRptMapper.getOrderNewQtyList(quarter, beginDate, endDate);
        Map<String, RPTCustomerReminderEntity> orderNewQtyMap = orderNewQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> reminderQtyList = customerReminderRptMapper.getReminderQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> reminderQtyMap = reminderQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> reminderFirstQtyList = customerReminderRptMapper.getReminderFirstQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> reminderFirstQtyMap = reminderFirstQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> reminderMultipleQtyList = customerReminderRptMapper.getReminderMultipleQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> reminderMultipleQtyMap = reminderMultipleQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> reminderOrderQtyList = customerReminderRptMapper.getReminderOrderQtyList(beginDate, endDate);
        Map<String, RPTCustomerReminderEntity> reminderOrderQtyMap = reminderOrderQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> exceed48hourReminderQtyList = customerReminderRptMapper.getExceed48hourReminderQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> exceed48hourReminderQtyMap = exceed48hourReminderQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> complete24hourQtyList = customerReminderRptMapper.getComplete24hourQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> complete24hourQtyMap = complete24hourQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTCustomerReminderEntity> over48ReminderCompletedQtyList = customerReminderRptMapper.getOver48ReminderCompletedQtyList(beginDt, endDt);
        Map<String, RPTCustomerReminderEntity> over48ReminderCompletedQtyMap = over48ReminderCompletedQtyList.stream().collect(Collectors.toMap(i -> i.getCustomerId() + ":" + i.getProductCategoryId(), Function.identity()));


        Set<String> keySet = Sets.newHashSet();
        keySet.addAll(orderNewQtyMap.keySet());
        keySet.addAll(reminderQtyMap.keySet());
        keySet.addAll(reminderFirstQtyMap.keySet());
        keySet.addAll(reminderMultipleQtyMap.keySet());
        keySet.addAll(reminderOrderQtyMap.keySet());
        keySet.addAll(complete24hourQtyMap.keySet());
        keySet.addAll(over48ReminderCompletedQtyMap.keySet());

        List<String> keyList = new ArrayList<>(keySet);
        RPTCustomerReminderEntity entity;
        RPTCustomerReminderEntity reminderEntity;
        Long productCategoryId;
        Long customerId;
        Integer orderNewQty, reminderQty, reminderFirstQty, reminderMultipleQty, reminderOrderQty, exceed48hourReminderQty, complete24hourQty, over48ReminderCompletedQty;
        int systemId = RptCommonUtils.getSystemId();
        for (String key : keyList) {
            entity = new RPTCustomerReminderEntity();
            entity.setSystemId(systemId);
            entity.setStatisticsDt(beginDt);
            entity.setQuarter(quarter);
            reminderEntity = orderNewQtyMap.get(key);
            if (reminderEntity != null) {
                orderNewQty = reminderEntity.getOrderNewQty();
                if (orderNewQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setOrderNewQty(orderNewQty);
                }
            }
            reminderEntity = reminderQtyMap.get(key);
            if (reminderEntity != null) {
                reminderQty = reminderEntity.getReminderQty();
                if (reminderQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setReminderQty(reminderQty);
                }
            }
            reminderEntity = reminderFirstQtyMap.get(key);
            if (reminderEntity != null) {
                reminderFirstQty = reminderEntity.getReminderFirstQty();
                if (reminderFirstQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setReminderFirstQty(reminderFirstQty);
                }
            }
            reminderEntity = reminderMultipleQtyMap.get(key);
            if (reminderEntity!= null) {
                reminderMultipleQty = reminderEntity.getReminderMultipleQty();
                if (reminderMultipleQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setReminderMultipleQty(reminderMultipleQty);
                }
            }
            reminderEntity = reminderOrderQtyMap.get(key);
            if (reminderEntity != null) {
                reminderOrderQty = reminderEntity.getReminderOrderQty();
                if (reminderOrderQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setReminderOrderQty(reminderOrderQty);
                }
            }
            reminderEntity = exceed48hourReminderQtyMap.get(key);
            if (reminderEntity != null) {
                exceed48hourReminderQty = reminderEntity.getExceed48hourReminderQty();
                if (exceed48hourReminderQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setExceed48hourReminderQty(exceed48hourReminderQty);
                }
            }
            reminderEntity = complete24hourQtyMap.get(key);
            if (reminderEntity != null) {
                complete24hourQty = reminderEntity.getComplete24hourQty();
                if (complete24hourQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setComplete24hourQty(complete24hourQty);
                }
            }
            reminderEntity = over48ReminderCompletedQtyMap.get(key);
            if (reminderEntity != null) {
                over48ReminderCompletedQty = reminderEntity.getOver48ReminderCompletedQty();
                if (over48ReminderCompletedQty != null) {
                    productCategoryId = reminderEntity.getProductCategoryId();
                    customerId = reminderEntity.getCustomerId();
                    if (productCategoryId != null) {
                        entity.setCustomerId(customerId);
                        entity.setProductCategoryId(productCategoryId);
                    }
                    entity.setOver48ReminderCompletedQty(over48ReminderCompletedQty);
                }
            }

            list.add(entity);
        }

        return list;
    }

    private Map<String, Long> getCustomerReminderIdMap(Integer systemId, Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(beginDate);
        List<RPTCustomerReminderEntity>  customerReminderIds = customerReminderRptMapper.getCustomerReminderIds(systemId, beginDate.getTime(), endDate.getTime(),quarter);
        if (customerReminderIds != null && !customerReminderIds.isEmpty()) {
            return customerReminderIds.stream().collect(Collectors.toMap(t->t.getStatisticsDt()+":"+t.getCustomerId()+":"+t.getProductCategoryId(), RPTCustomerReminderEntity::getId,(key1, key2) -> key2));
        } else {
            return Maps.newHashMap();
        }
    }

    public void saveCustomerReminderToRptDB(Date date) {
        List<RPTCustomerReminderEntity> list = getCustomerReminderListFromWebDB(date);
        if (!list.isEmpty()) {
            for (RPTCustomerReminderEntity entity : list) {
                customerReminderRptMapper.insertCustomerReminder(entity);
            }
        }
    }

    /**
     * 将工单系统中有的而中间表中没有的客户每日催单保存到中间表
     */
    public void saveMissedCustomerReminderToRptDB(Date date) {
        if (date != null) {
            List<RPTCustomerReminderEntity> list = getCustomerReminderListFromWebDB(date);
            if (!list.isEmpty()) {
                int systemId = RptCommonUtils.getSystemId();
                Map<String, Long> customerReminderIdMap = getCustomerReminderIdMap(systemId, date);
                Long primaryKeyId;
                for (RPTCustomerReminderEntity item : list) {
                    primaryKeyId = customerReminderIdMap.get(item.getStatisticsDt()+":"+item.getCustomerId()+":"+item.getProductCategoryId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        customerReminderRptMapper.insertCustomerReminder(item);
                    }
                }
            }
        }
    }

    /**
     * 删除中间表中指定日期的客户每日催单数据
     */
    public void deleteCustomerReminderFromRptDB(Date date) {
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            customerReminderRptMapper.deleteCustomerReminders(systemId, beginDate.getTime(), endDate.getTime());
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
                            saveCustomerReminderToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedCustomerReminderToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteCustomerReminderFromRptDB(beginDate);
                            saveCustomerReminderToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteCustomerReminderFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CustomerReminderRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerReminderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerReminderSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getCustomerId() != null && searchCondition.getSelectedYear() != null && searchCondition.getSelectedMonth() != null) {
            Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
            Date beginDate = DateUtils.getStartOfDay(queryDate);
            Date endDate = DateUtils.addMonth(beginDate, 1);
            String quarter = QuarterUtils.getSeasonQuarter(queryDate);
            Long customerId = searchCondition.getCustomerId();
            List<Long> productCategoryIds = searchCondition.getProductCategoryIds();
            long beginDateTime = beginDate.getTime();
            long endDateTime = endDate.getTime();
            int systemId = RptCommonUtils.getSystemId();
            Integer rowCount = customerReminderRptMapper.hasReportData(customerId,systemId,productCategoryIds,beginDateTime,endDateTime,quarter);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook customerReminderExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerReminderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerReminderSearch.class);
            List<RPTCustomerReminderEntity> list = getCustomerReminderList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(5000);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;


            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 12));

            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日接单量");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日催单数");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "占比");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "一次催单数");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "占比");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "多次催单数");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "占比");
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "超48时催单");
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "占比");
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "24小时完成率");
            ExportExcel.createCell(headFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "超时48时完成率");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            if (list != null && list.size() > 0) {
                NumberFormat numberFormat = NumberFormat.getInstance();
                numberFormat.setMaximumFractionDigits(2);
                Integer orderNewQty = 0, reminderQty = 0, reminderFirstQty = 0, reminderMultipleQty = 0,
                        reminderOrderQty = 0, exceed48hourReminderQty = 0, complete24hourQty = 0, over48ReminderCompletedQty = 0;
                for (int i = 0; i < list.size(); i++) {
                    RPTCustomerReminderEntity entity = list.get(i);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(entity.getStatisticsDate(), "yyyy-MM-dd"));
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCustomerName());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderNewQty());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderQty());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderRate() + "%");
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderFirstQty());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderFirstRate() + "%");
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderMultipleQty());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderMultipleRate() + "%");
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getExceed48hourReminderQty());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getExceed48hourReminderRate() + "%");
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getComplete24hourRate() + "%");
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOver48ReminderCompletedRate() + "%");
                    orderNewQty += entity.getOrderNewQty();
                    reminderQty += entity.getReminderQty();
                    reminderFirstQty += entity.getReminderFirstQty();
                    reminderMultipleQty += entity.getReminderMultipleQty();
                    reminderOrderQty += entity.getReminderOrderQty();
                    exceed48hourReminderQty += entity.getExceed48hourReminderQty();
                    complete24hourQty += entity.getComplete24hourQty();
                    over48ReminderCompletedQty += entity.getOver48ReminderCompletedQty();
                }

                Row sumRow = xSheet.createRow(rowIndex);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);


                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 1));
                ExportExcel.createCell(sumRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderNewQty);
                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, reminderQty);
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderNewQty == 0 ? "0.00" : numberFormat.format((float) reminderQty / orderNewQty * 100) + "%"));
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, reminderFirstQty);
                ExportExcel.createCell(sumRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderNewQty == 0 ? "0.00" : numberFormat.format((float) reminderFirstQty / orderNewQty * 100) + "%"));
                ExportExcel.createCell(sumRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, reminderMultipleQty);
                ExportExcel.createCell(sumRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderNewQty == 0 ? "0.00" : numberFormat.format((float) reminderMultipleQty / orderNewQty * 100) + "%"));
                ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, exceed48hourReminderQty);
                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderNewQty == 0 ? "0.00" : numberFormat.format((float) exceed48hourReminderQty / orderNewQty * 100) + "%"));
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (reminderOrderQty == 0 ? "0.00" : numberFormat.format((float) complete24hourQty / orderNewQty * 100) + "%"));
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (reminderOrderQty == 0 ? "0.00" : numberFormat.format((float) over48ReminderCompletedQty / orderNewQty * 100) + "%"));

            }
        }  catch (Exception e) {
            log.error("【CustomerReminderRptService.customerReminderExport】客户每日催单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
    }
        return xBook;
    }

}
