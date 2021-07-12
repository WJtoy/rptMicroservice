package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.common.RPTMiddleTableEnum;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CustomerChargeSummaryRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.utils.*;
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
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 客户对账单 - 工单数量与消费金额汇总
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerChargeSummaryRptService extends RptBaseService {

    @Resource
    private CustomerChargeSummaryRptMapper customerChargeSummaryRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private CompletedOrderRptService completedOrderRptService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;

    //region 公开

    public RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummary(RPTCustomerChargeSearch search) {
        RPTCustomerChargeSummaryMonthlyEntity result = null;
        if (search != null && search.getCustomerId() != null && search.getCustomerId() > 0
                && search.getSelectedYear() != null && search.getSelectedYear() > 0
                && search.getSelectedMonth() != null && search.getSelectedMonth() > 0) {
            int yearMonth = generateYearMonth(search.getSelectedYear(), search.getSelectedMonth());
            int currentYearMonth = generateYearMonth(new Date());

//            if (yearMonth == currentYearMonth) {
//                result = getCustomerChargeSummaryMonthlyByCurrentMonth(search.getCustomerId(), search.getSelectedYear(), search.getSelectedMonth());
//            } else
            if (yearMonth <= currentYearMonth) {
                result = getCustomerChargeSummaryMonthly(search.getCustomerId(), yearMonth);
            } else {
                result = new RPTCustomerChargeSummaryMonthlyEntity();
            }
        }
        return result;
    }

    public void saveCustomerChargeSummaryToRptDB(int selectedYear, int selectedMonth) {
        saveCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
        saveCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
    }

    public void updateCustomerChargeSummaryToRptDB(int selectedYear, int selectedMonth) {
        updateCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
        updateCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
    }

    public void deleteCustomerChargeSummaryFromRptDB(int selectedYear, int selectedMonth) {
        deleteCustomerOrderQtyMonthlysFromRptDB(selectedYear, selectedMonth);
        deleteCustomerFinanceMonthlysFromRptDB(selectedYear, selectedMonth);
    }

    //endregion 公开

    //region 从Web数据库获取客户的工单数量与消费金额数据

    /**
     * 查询所有客户的工单数量
     */
    public List<RPTCustomerChargeSummaryMonthlyEntity> getCustomerOrderQtyMonthlyListFromWebDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> result = Lists.newArrayList();
        Date startOfMonth = DateUtils.getStartOfDay(DateUtils.getDate(selectedYear, selectedMonth, 1));
        int yearMonth = generateYearMonth(startOfMonth);
        int currentYearMonth = generateYearMonth(new Date());
        Date endDate;
        if (yearMonth == currentYearMonth) {
            endDate = DateUtils.getStartOfDay(new Date());
        } else {
            endDate = DateUtils.addMonth(startOfMonth, 1);
        }
        String quarter = QuarterUtils.getSeasonQuarter(startOfMonth);

        List<RPTCustomerChargeSummaryMonthlyEntity> newQtyList = customerChargeSummaryRptMapper.getNewOrderQtyList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> completedQtyList = customerChargeSummaryRptMapper.getCompletedOrderQtyList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> returnedQtyList = customerChargeSummaryRptMapper.getReturnedOrderQtyList(startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> cancelledQtyList = customerChargeSummaryRptMapper.getCancelledOrderQtyList(startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> lastMonthUncompletedQtyList = customerChargeSummaryRptMapper.getUnCompletedOrderQtyList(startOfMonth);
        List<RPTCustomerChargeSummaryMonthlyEntity> uncompletedQtyList = customerChargeSummaryRptMapper.getUnCompletedOrderQtyList(endDate);

        Set<Long> customerIds = Sets.newHashSet();
        Map<Long, Integer> newQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : newQtyList) {
            customerIds.add(item.getCustomerId());
            newQtyMap.put(item.getCustomerId(), item.getNewQty());
        }
        Map<Long, Integer> completedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : completedQtyList) {
            customerIds.add(item.getCustomerId());
            completedQtyMap.put(item.getCustomerId(), item.getCompletedQty());
        }
        Map<Long, Integer> returnedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : returnedQtyList) {
            customerIds.add(item.getCustomerId());
            returnedQtyMap.put(item.getCustomerId(), item.getReturnedQty());
        }
        Map<Long, Integer> cancelledQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : cancelledQtyList) {
            customerIds.add(item.getCustomerId());
            cancelledQtyMap.put(item.getCustomerId(), item.getCancelledQty());
        }
        Map<Long, Integer> lastMonthUncompletedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : lastMonthUncompletedQtyList) {
            customerIds.add(item.getCustomerId());
            lastMonthUncompletedQtyMap.put(item.getCustomerId(), item.getUncompletedQty());
        }
        Map<Long, Integer> uncompletedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : uncompletedQtyList) {
            uncompletedQtyMap.put(item.getCustomerId(), item.getUncompletedQty());
            customerIds.add(item.getCustomerId());
        }
        RPTCustomerChargeSummaryMonthlyEntity entity;
        int yearmonth = generateYearMonth(startOfMonth);
        for (Long customerId : customerIds) {
            entity = new RPTCustomerChargeSummaryMonthlyEntity();
            entity.setCustomerId(customerId);
            entity.setYearmonth(yearmonth);
            entity.setNewQty(NumberUtils.toInteger(newQtyMap.get(customerId)));
            entity.setCompletedQty(NumberUtils.toInteger(completedQtyMap.get(customerId)));
            entity.setReturnedQty(NumberUtils.toInteger(returnedQtyMap.get(customerId)));
            entity.setCancelledQty(NumberUtils.toInteger(cancelledQtyMap.get(customerId)));
            entity.setLastMonthUncompletedQty(NumberUtils.toInteger(lastMonthUncompletedQtyMap.get(customerId)));
            entity.setUncompletedQty(NumberUtils.toInteger(uncompletedQtyMap.get(customerId)));
            result.add(entity);
        }
        return result;
    }


    /**
     * 查询所有客户的消费金额
     */
    private List<RPTCustomerChargeSummaryMonthlyEntity> getCustomerFinanceMonthlyListFromWebDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> result = Lists.newArrayList();
        Date startOfMonth = DateUtils.getStartOfDay(DateUtils.getDate(selectedYear, selectedMonth, 1));
        int yearMonth = generateYearMonth(startOfMonth);
        int currentYearMonth = generateYearMonth(new Date());
        Date endDate;
        if (yearMonth == currentYearMonth) {
            endDate = DateUtils.getStartOfDay(new Date());
        } else {
            endDate = DateUtils.addMonth(startOfMonth, 1);
        }
        String quarter = QuarterUtils.getSeasonQuarter(startOfMonth);

        List<RPTCustomerChargeSummaryMonthlyEntity> rechargeAmountList = customerChargeSummaryRptMapper.getRechargeAmountList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> completedOrderChargeList = customerChargeSummaryRptMapper.getCompletedOrderAndTimelinessAndUrgentChargeList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> writeOffChargeList = customerChargeSummaryRptMapper.getWriteOffChargeList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> lastMonthBalanceList = customerChargeSummaryRptMapper.getBalanceList(startOfMonth);
        List<RPTCustomerChargeSummaryMonthlyEntity> balanceList = customerChargeSummaryRptMapper.getBalanceList(endDate);
//        List<RPTCustomerChargeSummaryMonthlyEntity> blockAmountList = customerChargeSummaryRptMapper.getBlockAmountList(endDate);

        Set<Long> customerIds = Sets.newHashSet();
        Map<Long, Double> rechargeAmountMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : rechargeAmountList) {
            customerIds.add(item.getCustomerId());
            rechargeAmountMap.put(item.getCustomerId(), item.getRechargeAmount());
        }
        Map<Long, RPTCustomerChargeSummaryMonthlyEntity> completedOrderAndTimelinessAndUrgentChargeMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : completedOrderChargeList) {
            customerIds.add(item.getCustomerId());
            completedOrderAndTimelinessAndUrgentChargeMap.put(item.getCustomerId(), item);
        }
        Map<Long, Double> writeOffChargeMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : writeOffChargeList) {
            customerIds.add(item.getCustomerId());
            writeOffChargeMap.put(item.getCustomerId(), item.getWriteOffCharge());
        }
        Map<Long, Double> lastMonthBalanceMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : lastMonthBalanceList) {
            customerIds.add(item.getCustomerId());
            lastMonthBalanceMap.put(item.getCustomerId(), item.getBalance());
        }
        Map<Long, Double> balanceMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : balanceList) {
            balanceMap.put(item.getCustomerId(), item.getBalance());
            customerIds.add(item.getCustomerId());
        }
        Map<Long, Double> blockAmountMap = Maps.newHashMap();
//        for (RPTCustomerChargeSummaryMonthlyEntity item : blockAmountList) {
//            customerIds.add(item.getCustomerId());
//            blockAmountMap.put(item.getCustomerId(), item.getBlockAmount());
//        }
        RPTCustomerChargeSummaryMonthlyEntity entity;
        RPTCustomerChargeSummaryMonthlyEntity completedOrderCharge;
        int yearmonth = generateYearMonth(startOfMonth);
        for (Long customerId : customerIds) {
            entity = new RPTCustomerChargeSummaryMonthlyEntity();
            entity.setCustomerId(customerId);
            entity.setYearmonth(yearmonth);
            entity.setRechargeAmount(NumberUtils.toDouble(rechargeAmountMap.get(customerId)));
            completedOrderCharge = completedOrderAndTimelinessAndUrgentChargeMap.get(customerId);
            if (completedOrderCharge != null) {
                entity.setCompletedOrderCharge(NumberUtils.toDouble(completedOrderCharge.getCompletedOrderCharge()));
                entity.setTimelinessCharge(NumberUtils.toDouble(completedOrderCharge.getTimelinessCharge()));
                entity.setUrgentCharge(NumberUtils.toDouble(completedOrderCharge.getUrgentCharge()));
            }
            entity.setWriteOffCharge(NumberUtils.toDouble(writeOffChargeMap.get(customerId)));
            entity.setLastMonthBalance(NumberUtils.toDouble(lastMonthBalanceMap.get(customerId)));
            entity.setBalance(NumberUtils.toDouble(balanceMap.get(customerId)));
            entity.setBlockAmount(NumberUtils.toDouble(blockAmountMap.get(customerId)));
            result.add(entity);
        }
        return result;
    }


    /**
     * 查询单个客户的工单数量与消费金额
     */
    private RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryMonthlyByCurrentMonth(long customerId, int selectedYear, int selectedMonth) {
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        Date beginDate = DateUtils.getStartOfDay(queryDate);
        Date endDate = DateUtils.addMonth(beginDate, 1);
        Integer newQty = customerChargeSummaryRptMapper.getNewOrderQty(quarter, customerId, beginDate, endDate);
        Integer completedQty = customerChargeSummaryRptMapper.getCompletedOrderQty(quarter, customerId, beginDate, endDate);
        Integer returnedQty = customerChargeSummaryRptMapper.getReturnedOrderQty(customerId, beginDate, endDate);
        Integer cancelledQty = customerChargeSummaryRptMapper.getCancelledOrderQty(customerId, beginDate, endDate);

        Double rechargeAmount = customerChargeSummaryRptMapper.getRechargeAmount(quarter, customerId, beginDate, endDate);
        RPTCustomerChargeSummaryMonthlyEntity completedOrderAndTimelinessAndUrgentCharge = customerChargeSummaryRptMapper.getCompletedOrderAndTimelinessAndUrgentCharge(quarter, customerId, beginDate, endDate);
        Double writeOffCharge = customerChargeSummaryRptMapper.getWriteOffCharge(quarter, customerId, beginDate, endDate);
//        Double currentMonthBlockAmount = customerChargeSummaryRptMapper.getCurrentMonthBlockAmount(quarter, customerId, beginDate, endDate);

        int systemId = RptCommonUtils.getSystemId();
        Date lastMonth = DateUtils.addMonth(queryDate, -1);
        int lastYearMonth = generateYearMonth(lastMonth);
        RPTCustomerChargeSummaryMonthlyEntity lastMonthOrderQtyMonthly = customerChargeSummaryRptMapper.getCustomerOrderQtyMonthly(systemId, customerId, lastYearMonth);
        RPTCustomerChargeSummaryMonthlyEntity lastMonthFinanceMonthly = customerChargeSummaryRptMapper.getCustomerFinanceMonthly(systemId, customerId, lastYearMonth);

        RPTCustomerChargeSummaryMonthlyEntity result = new RPTCustomerChargeSummaryMonthlyEntity();
        result.setCustomerId(customerId);
        result.setYearmonth(generateYearMonth(selectedYear, selectedMonth));
        result.setLastMonthUncompletedQty(lastMonthOrderQtyMonthly == null ? 0 : NumberUtils.toInteger(lastMonthOrderQtyMonthly.getUncompletedQty()));
        result.setNewQty(NumberUtils.toInteger(newQty));
        result.setCompletedQty(NumberUtils.toInteger(completedQty));
        result.setReturnedQty(NumberUtils.toInteger(returnedQty));
        result.setCancelledQty(NumberUtils.toInteger(cancelledQty));
        result.setUncompletedQty(result.calcUncompletedQty());
        double lastMonthBlockAmount = 0;
        if (lastMonthFinanceMonthly != null) {
            result.setLastMonthBalance(NumberUtils.toDouble(lastMonthFinanceMonthly.getBalance()));
            lastMonthBlockAmount = NumberUtils.toDouble(lastMonthFinanceMonthly.getBlockAmount());
        }
        result.setRechargeAmount(NumberUtils.toDouble(rechargeAmount));
        if (completedOrderAndTimelinessAndUrgentCharge != null) {
            result.setCompletedOrderCharge(NumberUtils.toDouble(completedOrderAndTimelinessAndUrgentCharge.getCompletedOrderCharge()));
            result.setTimelinessCharge(NumberUtils.toDouble(completedOrderAndTimelinessAndUrgentCharge.getTimelinessCharge()));
            result.setUrgentCharge(NumberUtils.toDouble(completedOrderAndTimelinessAndUrgentCharge.getUrgentCharge()));
        }
        result.setWriteOffCharge(NumberUtils.toDouble(writeOffCharge));
        result.setBalance(result.calcBalance());
//        result.setBlockAmount(lastMonthBlockAmount + NumberUtils.toDouble(currentMonthBlockAmount));
        return result;
    }

    //endregion 从Web数据库获取客户的工单数量与消费金额数据

    //region 操作中间表

    private RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryMonthly(long customerId, int yearmonth) {
        int systemId = RptCommonUtils.getSystemId();
        RPTCustomerChargeSummaryMonthlyEntity orderQtyMonthly = customerChargeSummaryRptMapper.getCustomerOrderQtyMonthly(systemId, customerId, yearmonth);
        RPTCustomerChargeSummaryMonthlyEntity financeMonthly = customerChargeSummaryRptMapper.getCustomerFinanceMonthly(systemId, customerId, yearmonth);
        RPTCustomerChargeSummaryMonthlyEntity result = new RPTCustomerChargeSummaryMonthlyEntity();
        result.setCustomerId(customerId);
        result.setYearmonth(yearmonth);
        if (orderQtyMonthly != null) {
            result.setLastMonthUncompletedQty(orderQtyMonthly.getLastMonthUncompletedQty());
            result.setNewQty(orderQtyMonthly.getNewQty());
            result.setCompletedQty(orderQtyMonthly.getCompletedQty());
            result.setReturnedQty(orderQtyMonthly.getReturnedQty());
            result.setCancelledQty(orderQtyMonthly.getCancelledQty());
            result.setUncompletedQty(orderQtyMonthly.getUncompletedQty());
        }
        if (financeMonthly != null) {
            result.setLastMonthBalance(financeMonthly.getLastMonthBalance());
            result.setRechargeAmount(financeMonthly.getRechargeAmount());
            result.setCompletedOrderCharge(financeMonthly.getCompletedOrderCharge());
            result.setWriteOffCharge(financeMonthly.getWriteOffCharge());
            result.setTimelinessCharge(financeMonthly.getTimelinessCharge());
            result.setUrgentCharge(financeMonthly.getUrgentCharge());
            result.setBalance(financeMonthly.getBalance());
            result.setBlockAmount(financeMonthly.getBlockAmount());
        }
        return result;
    }

    private void saveCustomerOrderQtyMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerOrderQtyMonthlyListFromWebDB(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                item.setSystemId(systemId);
                item.setYearmonth(yearmonth);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(System.currentTimeMillis());
                customerChargeSummaryRptMapper.insertCustomerOrderQtyMonthly(item);
            }
        }
    }

    private void saveCustomerFinanceMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerFinanceMonthlyListFromWebDB(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                item.setSystemId(systemId);
                item.setYearmonth(yearmonth);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(System.currentTimeMillis());
                customerChargeSummaryRptMapper.insertCustomerFinanceMonthly(item);
            }
        }
    }

    private Map<Long, Long> getCustomerOrderQtyMonthlyIdMap(int systemId, int yearmonth) {
        List<LongTwoTuple> tuples = customerChargeSummaryRptMapper.getCustomerOrderQtyMonthlyIds(systemId, yearmonth);
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

    private Map<Long, Long> getCustomerFinanceMonthlyIdMap(int systemId, int yearmonth) {
        List<LongTwoTuple> tuples = customerChargeSummaryRptMapper.getCustomerFinanceMonthlyIds(systemId, yearmonth);
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

    private void updateCustomerOrderQtyMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerOrderQtyMonthlyListFromWebDB(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            Map<Long, Long> idMap = getCustomerOrderQtyMonthlyIdMap(systemId, yearmonth);
            Long primaryKeyId;
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                primaryKeyId = idMap.get(item.getCustomerId());
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(item.getCreateDt());
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    customerChargeSummaryRptMapper.updateCustomerOrderQtyMonthly(item);
                } else {
                    item.setSystemId(systemId);
                    item.setYearmonth(yearmonth);
                    customerChargeSummaryRptMapper.insertCustomerOrderQtyMonthly(item);
                }
            }
        }
    }

    private void updateCustomerFinanceMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerFinanceMonthlyListFromWebDB(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            Map<Long, Long> idMap = getCustomerFinanceMonthlyIdMap(systemId, yearmonth);
            Long primaryKeyId;
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                primaryKeyId = idMap.get(item.getCustomerId());
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(item.getCreateDt());
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    customerChargeSummaryRptMapper.updateCustomerFinanceMonthly(item);
                } else {
                    item.setSystemId(systemId);
                    item.setYearmonth(yearmonth);
                    customerChargeSummaryRptMapper.insertCustomerFinanceMonthly(item);
                }
            }
        }
    }

    private void deleteCustomerOrderQtyMonthlysFromRptDB(int selectedYear, int selectedMonth) {
        customerChargeSummaryRptMapper.deleteCustomerOrderQtyMonthly(RptCommonUtils.getSystemId(), generateYearMonth(selectedYear, selectedMonth));
    }

    private void deleteCustomerFinanceMonthlysFromRptDB(int selectedYear, int selectedMonth) {
        customerChargeSummaryRptMapper.deleteCustomerFinanceMonthly(RptCommonUtils.getSystemId(), generateYearMonth(selectedYear, selectedMonth));
    }

    //endregion

    //region 辅助方法

    private int generateYearMonth(Date date) {
        int selectedYear = DateUtils.getYear(date);
        int selectedMonth = DateUtils.getMonth(date);
        return generateYearMonth(selectedYear, selectedMonth);
    }

    private int generateYearMonth(int selectedYear, int selectedMonth) {
        return StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
    }

    //endregion 辅助方法


    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTMiddleTableEnum table, RPTRebuildOperationTypeEnum operationType, Integer selectedYear, Integer selectedMonth) {
        boolean result = false;
        if (table != null && operationType != null && selectedYear != null && selectedYear > 0
                && selectedMonth != null && selectedMonth > 0) {
            try {
                if (table == RPTMiddleTableEnum.RPT_CUSTOMER_ORDER_QTY_MONTHLY) {
                    switch (operationType) {
                        case INSERT:
                            saveCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case INSERT_MISSED_DATA:
                            updateCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case UPDATE:
                            deleteCustomerOrderQtyMonthlysFromRptDB(selectedYear, selectedMonth);
                            saveCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case DELETE:
                            deleteCustomerOrderQtyMonthlysFromRptDB(selectedYear, selectedMonth);
                            break;
                    }
                    result = true;
                } else if (table == RPTMiddleTableEnum.RPT_CUSTOMER_FINANCE_MONTHLY) {
                    switch (operationType) {
                        case INSERT:
                            saveCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case INSERT_MISSED_DATA:
                            updateCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case UPDATE:
                            deleteCustomerFinanceMonthlysFromRptDB(selectedYear, selectedMonth);
                            saveCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
                            break;
                        case DELETE:
                            deleteCustomerFinanceMonthlysFromRptDB(selectedYear, selectedMonth);
                            break;
                    }
                    result = true;
                }

            } catch (Exception e) {
                log.error("CustomerChargeSummaryRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 导出 客户对账单
     *
     * @return
     */
    public SXSSFWorkbook exportCustomerChargeRpt(String searchConditionJson, String reportTitle) {
        RPTCustomerChargeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerChargeSearch.class);
        RPTCustomerChargeSummaryMonthlyEntity item = getCustomerChargeSummary(searchCondition);
        SXSSFWorkbook xBook = null;
        try {
            long customerId = item.getCustomerId();

            RPTCustomer customer = msCustomerService.get(customerId);

            item.setCustomer(customer == null ? new RPTCustomer() : customer);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet("消费汇总");
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_20);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, item.getCustomer().getName() + "对账单");
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 7));


            Row firstRow = xSheet.createRow(rowIndex++);
            firstRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

            ExportExcel.createCell(firstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "提供服务公司：广东快可立服务有限公司");
            xSheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 6));

            CellRangeAddress region = new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 6);
            ExportExcel.setRegionBorder(region, xSheet, xBook);

            ExportExcel.createCell(firstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "所属账期:" + searchCondition.getSelectedYear() + "年" + searchCondition.getSelectedMonth() + "月");


            Row firsHeaderRow = xSheet.createRow(rowIndex++);
            firsHeaderRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(firsHeaderRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月未完成单");
            ExportExcel.createCell(firsHeaderRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月下单");
            ExportExcel.createCell(firsHeaderRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单");
            ExportExcel.createCell(firsHeaderRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退单");
            ExportExcel.createCell(firsHeaderRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月取消单");
            ExportExcel.createCell(firsHeaderRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月未完成单");
            ExportExcel.createCell(firsHeaderRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            ExportExcel.createCell(firsHeaderRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

            Row firstDataRow = xSheet.createRow(rowIndex++);
            firstDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(firstDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getLastMonthUncompletedQty());
            ExportExcel.createCell(firstDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getNewQty());
            ExportExcel.createCell(firstDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCompletedQty());
            ExportExcel.createCell(firstDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getReturnedQty());
            ExportExcel.createCell(firstDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCancelledQty());
            ExportExcel.createCell(firstDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUncompletedQty());
            ExportExcel.createCell(firstDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(firstDataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

            Row secondHeaderRow = xSheet.createRow(rowIndex++);
            secondHeaderRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondHeaderRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月消费余额");
            ExportExcel.createCell(secondHeaderRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月充值");
            ExportExcel.createCell(secondHeaderRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(secondHeaderRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "对帐差异单（本期退补款）");
            ExportExcel.createCell(secondHeaderRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月时效费");
            ExportExcel.createCell(secondHeaderRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月加急费");
            ExportExcel.createCell(secondHeaderRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月消费余额");
            ExportExcel.createCell(secondHeaderRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工冻结金额");

            Row secondDataRow = xSheet.createRow(rowIndex++);
            secondDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getLastMonthBalance());
            ExportExcel.createCell(secondDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getRechargeAmount());
            ExportExcel.createCell(secondDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCompletedOrderCharge());
            ExportExcel.createCell(secondDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getWriteOffCharge());
            ExportExcel.createCell(secondDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTimelinessCharge());
            ExportExcel.createCell(secondDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUrgentCharge());
            ExportExcel.createCell(secondDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBalance());
            ExportExcel.createCell(secondDataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBlockAmount());


            StringBuilder stringBuilder = new StringBuilder();

            Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
            String dayString = DateUtils.formatDate(DateUtils.getLastDayOfMonth(queryDate), "yyyy-MM-dd");
            stringBuilder.append("截止到" + dayString + "," + item.getCustomer().getName());
            stringBuilder.append(item.getBalance() > 0 ? "在广东快可立家电服务有限公司余额为" : "欠广东快可立家电服务有限公司服务款");

            double money = Math.abs(item.getBalance());
            String s = String.valueOf(money);
            CurrencyUtil nf = new CurrencyUtil(s);
            String bigMoney = nf.Convert();

            stringBuilder.append(String.format("%.2f", money) + "元 ");
            stringBuilder.append("（大写：");
            stringBuilder.append(bigMoney);
            stringBuilder.append("）");
            // 截止到2015年3月31号易品购商贸公司欠广东快可立家电服务有限公司服务款18445元（大写：壹万捌仟肆佰肆拾伍元正）

            rowIndex = rowIndex + 2;
            Row remarkRow = xSheet.createRow(rowIndex++);

            ExportExcel.createCell(remarkRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, stringBuilder.toString());
            xSheet.addMergedRegion(new CellRangeAddress(remarkRow.getRowNum(), remarkRow.getRowNum(), 0, 7));

            Row last1Row = xSheet.createRow(rowIndex++);
            ExportExcel.createCell(last1Row, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "确认（单位盖章）:");
            Row last2Row = xSheet.createRow(rowIndex++);
            ExportExcel.createCell(last2Row, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "日期:");

            //添加完工单Sheet
            List<RPTCompletedOrderEntity> customerCompletedOrdersRptData = completedOrderRptService.getCompletedOrderList(searchCondition);
            if (customerCompletedOrdersRptData.size() < 2000) {
                completedOrderRptService.addCustomerChargeCompleteRptSheet(xBook, xStyle, customerCompletedOrdersRptData);
            } else {
                completedOrderRptService.addCustomerChargeCompleteRptSheetMore2000(xBook, xStyle, customerCompletedOrdersRptData);
            }
            //添加退单/取消单Sheet
            List<RPTCancelledOrderEntity> returnedOrderDetail = cancelledOrderRptService.getCancelledOrder(searchCondition);

            if (returnedOrderDetail.size() < 2000) {
                cancelledOrderRptService.addCustomerChargeReturnCancelRptSheet(xBook, xStyle, returnedOrderDetail);
            } else {
                cancelledOrderRptService.addCustomerChargeReturnCancelRptSheetMore2000(xBook, xStyle, returnedOrderDetail);
            }

            //添加退补单Sheet
            List<RPTCustomerWriteOffEntity> WriteOffData = customerWriteOffRptService.getCustomerWriteOffList(searchCondition);
            if (WriteOffData.size() < 2000) {
                customerWriteOffRptService.addCustomerWriteOffRptSheet(xBook, xStyle, WriteOffData);
            } else {
                customerWriteOffRptService.addCustomerWriteOffRptSheetMore2000(xBook, xStyle, WriteOffData);
            }

        } catch (Exception e) {
            log.error("客户对账单写入excel失败");
        }
        return xBook;

    }
}
