package com.kkl.kklplus.provider.rpt.chart.service;

import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderQtyDailyChartEntity;
import com.kkl.kklplus.provider.rpt.chart.mapper.OrderQtyDailyChartMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class OrderQtyDailyChartService {

    @Resource
    private OrderQtyDailyChartMapper orderQtyDailyChartMapper;

    public Map<String, Object> getOrderQtyDailyChartData(RPTDataDrawingListSearch search) {
        Map<String, Object> map = new HashMap<>();
        int systemId = RptCommonUtils.getSystemId();
        Long startDate = DateUtils.getDateStart(new Date(search.getEndDate())).getTime();
        Long endDate = DateUtils.getDateEnd(new Date(search.getEndDate())).getTime();

        RPTOrderQtyDailyChartEntity entity = orderQtyDailyChartMapper.getOrderQtyDailyChartData(startDate, endDate, systemId);

        NumberFormat numberFormat = NumberFormat.getInstance();
        int completedOrderQty = 0;
        int cancelledOrderQty = 0;
        int autoCompletedQty = 0;
        int financialAuditQty = 0;
        int autoFinancialAuditQty = 0;
        int abnormalOrderQty = 0;
        int uncompletedOrderQty = 0;
        numberFormat.setMaximumFractionDigits(2);
        String autoCompletedRate = "0";
        if(entity != null){
            completedOrderQty = entity.getCompletedOrder();
            cancelledOrderQty = entity.getCancelledOrder();
            autoCompletedQty = entity.getAutoCompletedOrder();
            financialAuditQty = entity.getAutoFinancialAudit();
            autoFinancialAuditQty = entity.getAutoFinancialAudit();
            abnormalOrderQty = entity.getAbnormalAudit();
            uncompletedOrderQty = entity.getUncompletedOrder();
            if (completedOrderQty != 0) {
                autoCompletedRate = numberFormat.format((float) autoCompletedQty / completedOrderQty * 100);
            }
        }

        map.put("cancelledOrderQty", cancelledOrderQty);
        map.put("completedOrderQty", completedOrderQty);
        map.put("uncompletedOrderQty", uncompletedOrderQty);
        map.put("financialAuditQty", financialAuditQty);
        map.put("autoFinancialAuditQty", autoFinancialAuditQty);
        map.put("abnormalOrderQty", abnormalOrderQty);
        map.put("autoCompletedQty", autoCompletedQty);
        map.put("autoCompletedRate", autoCompletedRate);
        return map;
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
                            saveOrderQtyDailyToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteOrderQtyDailyFromRptDB(beginDate);
                            saveOrderQtyDailyToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteOrderQtyDailyFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("OrderQtyDailyStatisticsService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 删除中间表中指定日期的数据
     */
    private void deleteOrderQtyDailyFromRptDB(Date date) {
        if (date != null) {
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            orderQtyDailyChartMapper.deleteOrderQtyDailyFromRptDB(systemId, startDate.getTime(), endDate.getTime());
        }
    }

    public void saveOrderQtyDailyToRptDB(Date date) {

        int systemId = RptCommonUtils.getSystemId();
        Long startDate = DateUtils.getStartOfDay(date).getTime();
        Long endDate = DateUtils.getDateEnd(date).getTime();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        RPTOrderQtyDailyChartEntity entity = new RPTOrderQtyDailyChartEntity();

        Integer customerPlan = orderQtyDailyChartMapper.getPlanOrderData(startDate, endDate, systemId, quarter);

        Integer completedOrder = orderQtyDailyChartMapper.getCompletedOrderData(startDate, endDate, systemId, quarter);

        Integer cancelledOrder = orderQtyDailyChartMapper.getCancelledOrderData(startDate, endDate, systemId, quarter);

        Integer financialAudit = orderQtyDailyChartMapper.getFinancialAuditData(startDate, endDate, systemId, quarter);

        Integer autoCompleted = orderQtyDailyChartMapper.getAutoCompletedData(startDate, endDate, systemId, quarter);

        Integer autoFinancialAudit = orderQtyDailyChartMapper.getAutoFinancialAuditData(startDate, endDate, systemId);

        Integer abnormalOrder = orderQtyDailyChartMapper.getAbnormalOrderData(startDate, endDate, systemId);


        long lastStartDate = DateUtils.addDays(new Date(startDate), -1).getTime();
        long lastEndDate = DateUtils.addDays(new Date(endDate), -1).getTime();

        Integer lastUnCompletedOrder = orderQtyDailyChartMapper.getUnCompletedOrderData(lastStartDate, lastEndDate, systemId);

        int unCompletedOrder = 0;
        if (customerPlan != null) {
            if (completedOrder == null) {
                completedOrder = 0;
            }
            if (cancelledOrder == null) {
                cancelledOrder = 0;
            }
            unCompletedOrder = customerPlan - completedOrder - cancelledOrder;
        }

        if (lastUnCompletedOrder == null) {
            lastUnCompletedOrder = 0;
        }
        entity.setSystemId(systemId);
        entity.setCreateDate(startDate);
        entity.setCustomerPlan(customerPlan == null ? 0 : customerPlan);
        entity.setCompletedOrder(completedOrder);
        entity.setCancelledOrder(cancelledOrder);
        entity.setFinancialAudit(financialAudit == null ? 0 : financialAudit);
        entity.setUncompletedOrder(lastUnCompletedOrder + unCompletedOrder);
        entity.setAutoCompletedOrder(autoCompleted == null ? 0 : autoCompleted);
        entity.setAutoFinancialAudit(autoFinancialAudit == null ? 0 : autoFinancialAudit);
        entity.setAbnormalAudit(abnormalOrder == null ? 0 : abnormalOrder);

        orderQtyDailyChartMapper.insertOrderQtyDailyData(entity);

    }

}
