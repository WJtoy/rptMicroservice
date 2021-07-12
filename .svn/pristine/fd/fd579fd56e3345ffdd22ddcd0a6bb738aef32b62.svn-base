package com.kkl.kklplus.provider.rpt.chart.service;

import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderCrushQtyEntity;
import com.kkl.kklplus.provider.rpt.chart.mapper.OrderCrushQtyChartMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class OrderCrushQtyChartService {

    @Resource
    private OrderCrushQtyChartMapper orderCrushQtyChartMapper;

    public Map<String,Object> getOrderCrushQtyData(RPTDataDrawingListSearch search){

        Map<String, Object> map = new HashMap<>();
        int systemId = RptCommonUtils.getSystemId();
        Long endDate = DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime();
        Long startDate = DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime();
        RPTOrderCrushQtyEntity entity = orderCrushQtyChartMapper.getOrderCrushQtyData(systemId,startDate,endDate);

        Integer plainOrderQty = orderCrushQtyChartMapper.getPlainOrderQty(systemId,startDate,endDate);

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        int orderCrushQty = 0;
        int onceCrushQty = 0;
        int repeatedlyCrushQty = 0;
        int completedCrushQty = 0;
        int completedOnceCrush = 0;
        int completedRepeatedlyCrush = 0;

        int completed24hour = 0;
        int completed48hour = 0;
        int over48hourCompleted = 0;

        if(entity != null){
            orderCrushQty = entity.getOrderCrushQty();
            onceCrushQty = entity.getOnceCrushQty();
            repeatedlyCrushQty = entity.getRepeatedlyCrushQty();
            completedCrushQty = entity.getCompletedCrushQty();
            completedOnceCrush = entity.getCompletedOnceCrush();
            completedRepeatedlyCrush = entity.getCompletedRepeatedlyCrush();

            completed24hour = entity.getCompleted24hour();
            completed48hour = entity.getCompleted48hour();
            over48hourCompleted = entity.getOver48hourCompleted();
        }

        String completedOnceCrushRate = "0";
        String completedRepeatedlyCrushRate = "0";

        String onceCrushPlainOrderRate = "0";

        if(completedCrushQty != 0){
            completedOnceCrushRate = numberFormat.format((float) completedOnceCrush / completedCrushQty * 100);
            completedRepeatedlyCrushRate = numberFormat.format((float) completedRepeatedlyCrush / completedCrushQty * 100);
        }
        if(plainOrderQty != null && plainOrderQty != 0){
            onceCrushPlainOrderRate = numberFormat.format((float) onceCrushQty / plainOrderQty * 100);
        }
        map.put("completedOnceCrushRate",completedOnceCrushRate);
        map.put("completedRepeatedlyCrushRate",completedRepeatedlyCrushRate);

        map.put("orderCrushQty",orderCrushQty);
        map.put("onceCrushQty",onceCrushQty);
        map.put("repeatedlyCrushQty",repeatedlyCrushQty);
        map.put("completedCrushQty",completedCrushQty);
        map.put("completedOnceCrush",completedOnceCrush);
        map.put("completedRepeatedlyCrush",completedRepeatedlyCrush);
        map.put("completed24hour",completed24hour);
        map.put("completed48hour",completed48hour);
        map.put("over48hourCompleted",over48hourCompleted);

        map.put("onceCrushPlainOrderRate",onceCrushPlainOrderRate);
        return map;
    }

    public void saveOrderCrushQtyToRptDB(Date date) {
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        int systemId = RptCommonUtils.getSystemId();
        RPTOrderCrushQtyEntity entity = new RPTOrderCrushQtyEntity();
        Integer orderCrushQty = orderCrushQtyChartMapper.getOrderCrushQty(startDate, endDate);
        Integer onceCrushQty = orderCrushQtyChartMapper.getOnceCrushQty(startDate, endDate);
        Integer repeatedlyCrushQty = orderCrushQtyChartMapper.getRepeatedlyCrushQty(startDate, endDate);
        Integer completedCrushQty = orderCrushQtyChartMapper.getCompletedCrushQty(startDate, endDate);
        Integer completedOnceCrush = orderCrushQtyChartMapper.getCompletedOnceCrush(startDate, endDate);
        Integer completedRepeatedlyCrush = orderCrushQtyChartMapper.getCompletedRepeatedlyCrush(startDate, endDate);

        List<RPTOrderCrushQtyEntity> completedOrderCrush = orderCrushQtyChartMapper.getCompletedOrderCrush(startDate, endDate);
        double timeDifference;
        int completed24hourQty = 0;
        int completed48hourQty = 0;
        int completedOver48hourQty = 0;
        for (RPTOrderCrushQtyEntity crushQtyEntity : completedOrderCrush) {
            timeDifference = DateUtils.getDateDiffHour(crushQtyEntity.getCrushCreateDate(), crushQtyEntity.getOrderCloseDate());
            if (timeDifference <= 24) {
                completed24hourQty = completed24hourQty + 1;
            } else if (timeDifference <= 48) {
                completed48hourQty = completed48hourQty + 1;
            } else {
                completedOver48hourQty = completedOver48hourQty + 1;
            }
        }
        entity.setSystemId(systemId);
        entity.setCreateDate(startDate.getTime());
        entity.setOrderCrushQty(orderCrushQty);
        entity.setOnceCrushQty(onceCrushQty);
        entity.setRepeatedlyCrushQty(repeatedlyCrushQty);
        entity.setCompletedCrushQty(completedCrushQty);
        entity.setCompletedOnceCrush(completedOnceCrush);
        entity.setCompletedRepeatedlyCrush(completedRepeatedlyCrush);
        entity.setCompleted24hour(completed24hourQty);
        entity.setCompleted48hour(completed48hourQty);
        entity.setOver48hourCompleted(completedOver48hourQty);

        orderCrushQtyChartMapper.insertOrderCrushQtyData(entity);
    }
    /**
     * 删除中间表中指定日期的数据
     */
    private void deleteOrderCrushQtyFromRptDB(Date date) {
        if (date != null) {
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            orderCrushQtyChartMapper.deleteOrderCrushQtyFromRptDB(systemId, startDate.getTime(), endDate.getTime());
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
                            saveOrderCrushQtyToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteOrderCrushQtyFromRptDB(beginDate);
                            saveOrderCrushQtyToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteOrderCrushQtyFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("OrderCrushQtyChartService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

}
