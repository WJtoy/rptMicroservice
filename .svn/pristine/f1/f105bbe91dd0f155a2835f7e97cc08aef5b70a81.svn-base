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
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CustomerChargeSummaryRptMapper;
import com.kkl.kklplus.provider.rpt.mapper.CustomerChargeSummaryRptNewMapper;
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
public class CustomerChargeSummaryRptNewService extends RptBaseService {

    @Resource
    private CustomerChargeSummaryRptNewMapper customerChargeSummaryRptNewMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private CompletedOrderRptService completedOrderRptService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;



    //endregion 公开

    public RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryNew(RPTCustomerChargeSearch search) {
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

    public void updateCustomerChargeSummaryToRptDB(int selectedYear, int selectedMonth) {
        updateCustomerOrderQtyMonthlysToRptDB(selectedYear, selectedMonth);
        updateCustomerFinanceMonthlysToRptDB(selectedYear, selectedMonth);
    }

    //region 从Web数据库获取客户的工单数量与消费金额数据

    /**
     * 查询所有客户的工单数量
     */
    public List<RPTCustomerChargeSummaryMonthlyEntity> getCustomerOrderQtyMonthlyListFromWebDBNew(int selectedYear, int selectedMonth) {
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

        List<RPTCustomerChargeSummaryMonthlyEntity> newQtyList = customerChargeSummaryRptNewMapper.getNewOrderQtyList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> completedQtyList = customerChargeSummaryRptNewMapper.getCompletedOrderQtyList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> returnedQtyList = customerChargeSummaryRptNewMapper.getReturnedOrderQtyList(startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> cancelledQtyList = customerChargeSummaryRptNewMapper.getCancelledOrderQtyList(startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> lastMonthUncompletedQtyList = customerChargeSummaryRptNewMapper.getUnCompletedOrderQtyList(startOfMonth);
        List<RPTCustomerChargeSummaryMonthlyEntity> uncompletedQtyList = customerChargeSummaryRptNewMapper.getUnCompletedOrderQtyList(endDate);

        Set<Long> customerIds = Sets.newHashSet();
        Set<String> keys = Sets.newHashSet();
        Map<String, Integer> newQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : newQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            customerIds.add(item.getCustomerId());
            newQtyMap.put(key, item.getNewQty());
        }
        Map<String, Integer> completedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : completedQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            customerIds.add(item.getCustomerId());
            completedQtyMap.put(key, item.getCompletedQty());
        }
        Map<String, Integer> returnedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : returnedQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            customerIds.add(item.getCustomerId());
            returnedQtyMap.put(key, item.getReturnedQty());
        }
        Map<String, Integer> cancelledQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : cancelledQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            customerIds.add(item.getCustomerId());
            cancelledQtyMap.put(key, item.getCancelledQty());
        }
        Map<String, Integer> lastMonthUncompletedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : lastMonthUncompletedQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            customerIds.add(item.getCustomerId());
            lastMonthUncompletedQtyMap.put(key, item.getUncompletedQty());
        }
        Map<String, Integer> uncompletedQtyMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : uncompletedQtyList) {
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            uncompletedQtyMap.put(key, item.getUncompletedQty());
            customerIds.add(item.getCustomerId());
            keys.add(key);
        }
        RPTCustomerChargeSummaryMonthlyEntity entity;
        int yearmonth = generateYearMonth(startOfMonth);
        for (String key : keys) {
            entity = new RPTCustomerChargeSummaryMonthlyEntity();
            String[] arr = key.split("%");
            entity.setCustomerId(Long.valueOf(arr[0]));
            entity.setProductCategoryId(Long.valueOf(arr[1]));
            entity.setYearmonth(yearmonth);
            entity.setNewQty(NumberUtils.toInteger(newQtyMap.get(key)));
            entity.setCompletedQty(NumberUtils.toInteger(completedQtyMap.get(key)));
            entity.setReturnedQty(NumberUtils.toInteger(returnedQtyMap.get(key)));
            entity.setCancelledQty(NumberUtils.toInteger(cancelledQtyMap.get(key)));
            entity.setLastMonthUncompletedQty(NumberUtils.toInteger(lastMonthUncompletedQtyMap.get(key)));
            entity.setUncompletedQty(NumberUtils.toInteger(uncompletedQtyMap.get(key)));
            result.add(entity);
        }

        Map<Long, List<RPTCustomerChargeSummaryMonthlyEntity>> orderDailyMap = result.stream().collect(Collectors.groupingBy(RPTCustomerChargeSummaryMonthlyEntity::getCustomerId));

        for(long customerId : customerIds){
            List<RPTCustomerChargeSummaryMonthlyEntity> list = orderDailyMap.get(customerId);
            RPTCustomerChargeSummaryMonthlyEntity customerEntity = new RPTCustomerChargeSummaryMonthlyEntity();
            int lastMonthUncompletedQty = 0;
            int newQty = 0;
            int completedQty = 0;
            int returnQty = 0;
            int cancelQty = 0;
            int uncompletedQty = 0;
            for(RPTCustomerChargeSummaryMonthlyEntity  item : list){
                newQty = newQty + item.getNewQty();
                completedQty = completedQty + item.getCompletedQty();
                returnQty = returnQty + item.getReturnedQty();
                cancelQty = cancelQty + item.getCancelledQty();
                uncompletedQty = uncompletedQty + item.getUncompletedQty();
                lastMonthUncompletedQty = lastMonthUncompletedQty + item.getLastMonthUncompletedQty();
            }
            customerEntity.setNewQty(newQty);
            customerEntity.setCompletedQty(completedQty);
            customerEntity.setReturnedQty(returnQty);
            customerEntity.setCancelledQty(cancelQty);
            customerEntity.setUncompletedQty(uncompletedQty);
            customerEntity.setLastMonthUncompletedQty(lastMonthUncompletedQty);
            customerEntity.setCustomerId(customerId);
            customerEntity.setProductCategoryId(0L);
            customerEntity.setYearmonth(yearmonth);
            result.add(customerEntity);
        }
        return result;
    }


    /**
     * 查询所有客户的消费金额
     */
    private List<RPTCustomerChargeSummaryMonthlyEntity> getCustomerFinanceMonthlyListFromWebDBNew(int selectedYear, int selectedMonth) {
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

        List<RPTCustomerChargeSummaryMonthlyEntity> rechargeAmountList = customerChargeSummaryRptNewMapper.getRechargeAmountList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> completedOrderChargeList = customerChargeSummaryRptNewMapper.getCompletedOrderAndTimelinessAndUrgentChargeList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> writeOffChargeList = customerChargeSummaryRptNewMapper.getWriteOffChargeList(quarter, startOfMonth, endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> lastMonthBalanceList = customerChargeSummaryRptNewMapper.getBalanceList(startOfMonth);
        List<RPTCustomerChargeSummaryMonthlyEntity> balanceList = customerChargeSummaryRptNewMapper.getBalanceList(endDate);
        List<RPTCustomerChargeSummaryMonthlyEntity> blockAmountList = customerChargeSummaryRptNewMapper.getBlockAmountList(endDate);

        Set<Long> customerIds = Sets.newHashSet();
        Set<String> keys = Sets.newHashSet();
        Map<Long, Double> rechargeAmountMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : rechargeAmountList) {
            customerIds.add(item.getCustomerId());
            rechargeAmountMap.put(item.getCustomerId(), item.getRechargeAmount());
        }
        Map<String, RPTCustomerChargeSummaryMonthlyEntity> completedOrderAndTimelinessAndUrgentChargeMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : completedOrderChargeList) {
            customerIds.add(item.getCustomerId());
            String key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            completedOrderAndTimelinessAndUrgentChargeMap.put(key, item);
        }
        Map<String, RPTCustomerChargeSummaryMonthlyEntity> writeOffChargeMap = Maps.newHashMap();
        for (RPTCustomerChargeSummaryMonthlyEntity item : writeOffChargeList) {
            customerIds.add(item.getCustomerId());
            String key =  StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
            keys.add(key);
            writeOffChargeMap.put(key, item);
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
        for (RPTCustomerChargeSummaryMonthlyEntity item : blockAmountList) {
            customerIds.add(item.getCustomerId());
            blockAmountMap.put(item.getCustomerId(), item.getBlockAmount());
        }
        RPTCustomerChargeSummaryMonthlyEntity entity;
        RPTCustomerChargeSummaryMonthlyEntity completedOrderCharge;
        RPTCustomerChargeSummaryMonthlyEntity writeOffEntity;
        double completeds;
        double writes;
        int yearmonth = generateYearMonth(startOfMonth);
        for (String key : keys) {
            completeds = 0.00;
            writes = 0.00;
            entity = new RPTCustomerChargeSummaryMonthlyEntity();
            String[] arr = key.split("%");
            entity.setCustomerId(Long.valueOf(arr[0]));
            entity.setProductCategoryId(Long.valueOf(arr[1]));
            entity.setYearmonth(yearmonth);
            completedOrderCharge = completedOrderAndTimelinessAndUrgentChargeMap.get(key);
            if (completedOrderCharge != null) {
                entity.setCompletedOrderCharge(NumberUtils.toDouble(completedOrderCharge.getCompletedOrderCharge()));
                entity.setTimelinessCharge(NumberUtils.toDouble(completedOrderCharge.getTimelinessCharge()));
                entity.setUrgentCharge(NumberUtils.toDouble(completedOrderCharge.getUrgentCharge()));
                completeds = NumberUtils.toDouble(completedOrderCharge.getPraiseFee());
            }
            writeOffEntity = writeOffChargeMap.get(key);
            if (writeOffEntity != null) {
             entity.setWriteOffCharge(NumberUtils.toDouble(writeOffEntity.getWriteOffCharge()));
             writes =  NumberUtils.toDouble(writeOffEntity.getPraiseFee());
            }
            entity.setPraiseFee(writes + completeds);
            result.add(entity);
        }

        Map<Long, List<RPTCustomerChargeSummaryMonthlyEntity>> orderMap = result.stream().collect(Collectors.groupingBy(RPTCustomerChargeSummaryMonthlyEntity::getCustomerId));

        for(Long customerId :customerIds){
            List<RPTCustomerChargeSummaryMonthlyEntity> list = orderMap.get(customerId);
            RPTCustomerChargeSummaryMonthlyEntity customerEntity = new RPTCustomerChargeSummaryMonthlyEntity();

            double completedCharge = 0.00;
            double timelinessCharge = 0.00;
            double urgentCharge = 0.00;
            double writeOffCharge = 0.00;
            double praiseFee = 0.00;
            if(list !=null){
                for(RPTCustomerChargeSummaryMonthlyEntity item : list){
                     completedCharge = completedCharge + item.getCompletedOrderCharge();
                     timelinessCharge = timelinessCharge + item.getTimelinessCharge();
                     urgentCharge = urgentCharge + item.getUrgentCharge();
                     writeOffCharge = writeOffCharge + item.getWriteOffCharge();
                     praiseFee =  praiseFee + item.getPraiseFee();
                }
            }
            customerEntity.setRechargeAmount(NumberUtils.toDouble(rechargeAmountMap.get(customerId)));
            customerEntity.setLastMonthBalance(NumberUtils.toDouble(lastMonthBalanceMap.get(customerId)));
            customerEntity.setBalance(NumberUtils.toDouble(balanceMap.get(customerId)));
            customerEntity.setBlockAmount(NumberUtils.toDouble(blockAmountMap.get(customerId)));
            customerEntity.setCompletedOrderCharge(completedCharge);
            customerEntity.setTimelinessCharge(timelinessCharge);
            customerEntity.setUrgentCharge(urgentCharge);
            customerEntity.setWriteOffCharge(writeOffCharge);
            customerEntity.setPraiseFee(praiseFee);
            customerEntity.setCustomerId(customerId);
            customerEntity.setProductCategoryId(0L);
            customerEntity.setYearmonth(yearmonth);
            result.add(customerEntity);
        }
        return result;
    }

    private RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryMonthly(long customerId, int yearmonth) {
        int systemId = RptCommonUtils.getSystemId();
        RPTCustomerChargeSummaryMonthlyEntity orderQtyMonthly = customerChargeSummaryRptNewMapper.getCustomerOrderQtyMonthly(systemId, customerId, yearmonth);
        RPTCustomerChargeSummaryMonthlyEntity financeMonthly = customerChargeSummaryRptNewMapper.getCustomerFinanceMonthly(systemId, customerId, yearmonth);
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
            result.setPraiseFee(financeMonthly.getPraiseFee());
            result.setBalance(financeMonthly.getBalance());
            result.setBlockAmount(financeMonthly.getBlockAmount());
        }
        return result;
    }




    //endregion 从Web数据库获取客户的工单数量与消费金额数据

    //region 操作中间表

    private void saveCustomerOrderQtyMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerOrderQtyMonthlyListFromWebDBNew(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                item.setSystemId(systemId);
                item.setYearmonth(yearmonth);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(System.currentTimeMillis());
                customerChargeSummaryRptNewMapper.insertCustomerOrderQtyMonthly(item);
            }
        }
    }

    private void saveCustomerFinanceMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerFinanceMonthlyListFromWebDBNew(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                item.setSystemId(systemId);
                item.setYearmonth(yearmonth);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(System.currentTimeMillis());
                customerChargeSummaryRptNewMapper.insertCustomerFinanceMonthly(item);
            }
        }
    }

    private Map<String, Long> getCustomerOrderQtyMonthlyIdMap(int systemId, int yearmonth) {
        List<LongThreeTuple> tuples = customerChargeSummaryRptNewMapper.getCustomerOrderQtyMonthlyIds(systemId, yearmonth);
        Map<String, Long> tuplesMap = Maps.newHashMap();
        if (tuples != null && !tuples.isEmpty()) {
            for(LongThreeTuple item : tuples ){
                String key = StringUtils.join(item.getBElement(), "%", item.getCElement());
                tuplesMap.put(key,item.getAElement());
            }
            return tuplesMap;
        } else {
            return tuplesMap;
        }
    }
    private Map<String, Long> getCustomerFinanceMonthlyIdMap(int systemId, int yearmonth) {
        List<LongThreeTuple> tuples = customerChargeSummaryRptNewMapper.getCustomerFinanceMonthlyIds(systemId, yearmonth);
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

    private void updateCustomerOrderQtyMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerOrderQtyMonthlyListFromWebDBNew(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            Map<String, Long> idMap = getCustomerOrderQtyMonthlyIdMap(systemId, yearmonth);
            Long primaryKeyId;
            String key;
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
                primaryKeyId = idMap.get(key);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(item.getCreateDt());
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    customerChargeSummaryRptNewMapper.updateCustomerOrderQtyMonthly(item);
                } else {
                    item.setSystemId(systemId);
                    item.setYearmonth(yearmonth);
                    customerChargeSummaryRptNewMapper.insertCustomerOrderQtyMonthly(item);
                }
            }
        }
    }

    private void updateCustomerFinanceMonthlysToRptDB(int selectedYear, int selectedMonth) {
        List<RPTCustomerChargeSummaryMonthlyEntity> list = getCustomerFinanceMonthlyListFromWebDBNew(selectedYear, selectedMonth);
        if (!list.isEmpty()) {
            int yearmonth = generateYearMonth(selectedYear, selectedMonth);
            int systemId = RptCommonUtils.getSystemId();
            Map<String, Long> idMap = getCustomerFinanceMonthlyIdMap(systemId, yearmonth);
            Long primaryKeyId;
            String key;
            for (RPTCustomerChargeSummaryMonthlyEntity item : list) {
                key = StringUtils.join(item.getCustomerId(), "%", item.getProductCategoryId());
                primaryKeyId = idMap.get(key);
                item.setCreateDt(System.currentTimeMillis());
                item.setUpdateDt(item.getCreateDt());
                if (primaryKeyId != null && primaryKeyId != 0) {
                    item.setId(primaryKeyId);
                    customerChargeSummaryRptNewMapper.updateCustomerFinanceMonthly(item);
                } else {
                    item.setSystemId(systemId);
                    item.setYearmonth(yearmonth);
                    customerChargeSummaryRptNewMapper.insertCustomerFinanceMonthly(item);
                }
            }
        }
    }

    private void deleteCustomerOrderQtyMonthlysFromRptDB(int selectedYear, int selectedMonth) {
        customerChargeSummaryRptNewMapper.deleteCustomerOrderQtyMonthly(RptCommonUtils.getSystemId(), generateYearMonth(selectedYear, selectedMonth));
    }

    private void deleteCustomerFinanceMonthlysFromRptDB(int selectedYear, int selectedMonth) {
        customerChargeSummaryRptNewMapper.deleteCustomerFinanceMonthly(RptCommonUtils.getSystemId(), generateYearMonth(selectedYear, selectedMonth));
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
    public SXSSFWorkbook exportCustomerChargeRptNew(String searchConditionJson, String reportTitle) {
        RPTCustomerChargeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerChargeSearch.class);
        RPTCustomerChargeSummaryMonthlyEntity item = getCustomerChargeSummaryNew(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 8));


            Row firstRow = xSheet.createRow(rowIndex++);
            firstRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

            ExportExcel.createCell(firstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "提供服务公司：广东快可立服务有限公司");
            xSheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 7));

            CellRangeAddress region = new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 7);
            ExportExcel.setRegionBorder(region, xSheet, xBook);

            ExportExcel.createCell(firstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "所属账期:" + searchCondition.getSelectedYear() + "年" + searchCondition.getSelectedMonth() + "月");


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
            ExportExcel.createCell(firsHeaderRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

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
            ExportExcel.createCell(firstDataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

            Row secondHeaderRow = xSheet.createRow(rowIndex++);
            secondHeaderRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondHeaderRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月消费余额");
            ExportExcel.createCell(secondHeaderRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月充值");
            ExportExcel.createCell(secondHeaderRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(secondHeaderRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "对帐差异单（本期退补款）");
            ExportExcel.createCell(secondHeaderRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月时效费");
            ExportExcel.createCell(secondHeaderRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月加急费");
            ExportExcel.createCell(secondHeaderRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月好评费");
            ExportExcel.createCell(secondHeaderRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月消费余额");
            ExportExcel.createCell(secondHeaderRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工冻结金额");

            Row secondDataRow = xSheet.createRow(rowIndex++);
            secondDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getLastMonthBalance());
            ExportExcel.createCell(secondDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getRechargeAmount());
            ExportExcel.createCell(secondDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCompletedOrderCharge());
            ExportExcel.createCell(secondDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getWriteOffCharge());
            ExportExcel.createCell(secondDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTimelinessCharge());
            ExportExcel.createCell(secondDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUrgentCharge());
            ExportExcel.createCell(secondDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getPraiseFee());
            ExportExcel.createCell(secondDataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBalance());
            ExportExcel.createCell(secondDataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBlockAmount());


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
            ExportExcel.createCell(last1Row, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "确认（单位盖章）:");
            Row last2Row = xSheet.createRow(rowIndex++);
            ExportExcel.createCell(last2Row, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "日期:");


            List<RPTCompletedOrderEntity> customerCompletedOrdersRptData = completedOrderRptService.getCompletedOrderList(searchCondition);
            List<RPTCancelledOrderEntity> returnedOrderDetail = cancelledOrderRptService.getCancelledOrder(searchCondition);
            List<RPTCustomerWriteOffEntity> WriteOffData = customerWriteOffRptService.getCustomerWriteOffList(searchCondition);
            //添加完工单Sheet
            if (customerCompletedOrdersRptData.size() < 2000) {
                completedOrderRptService.addCustomerChargeCompleteRptSheet(xBook, xStyle, customerCompletedOrdersRptData);
            } else {
                completedOrderRptService.addCustomerChargeCompleteRptSheetMore2000(xBook, xStyle, customerCompletedOrdersRptData);
            }
            //添加退单/取消单Sheet
            if (returnedOrderDetail.size() < 2000) {
                cancelledOrderRptService.addCustomerChargeReturnCancelRptSheet(xBook, xStyle, returnedOrderDetail);
            } else {
                cancelledOrderRptService.addCustomerChargeReturnCancelRptSheetMore2000(xBook, xStyle, returnedOrderDetail);
            }

            //添加退补单Sheet
            if (WriteOffData.size() < 2000) {
                customerWriteOffRptService.addCustomerWriteOffRptSheet(xBook, xStyle, WriteOffData);
            } else {
                customerWriteOffRptService.addCustomerWriteOffRptSheetMore2000(xBook, xStyle, WriteOffData);
            }

        } catch (Exception e) {
            log.error("【CustomerChargeSummaryRptNewService.exportCustomerChargeRptNew】客户对账单写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

}
