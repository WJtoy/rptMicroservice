package com.kkl.kklplus.provider.rpt.chart.service;

import com.kkl.kklplus.entity.rpt.RPTCustomerReminderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.provider.rpt.chart.mapper.CustomerReminderChartMapper;
import com.kkl.kklplus.provider.rpt.chart.mapper.KeFuCompleteTimeChartMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerReminderChartService {

    @Resource
    private CustomerReminderChartMapper customerReminderChartMapper;

    /**
     * 获取每日催单汇总图表数据
     *
     * @param search
     * @return
     */
    public RPTCustomerReminderEntity getCustomerReminderChartData(RPTDataDrawingListSearch search) {
        int systemId = RptCommonUtils.getSystemId();
        Long endDate = DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime();
        Long startDate = DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime();
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        RPTCustomerReminderEntity entity = customerReminderChartMapper.getCustomerReminderList(systemId, startDate, endDate, quarter);
        if(entity == null){
            return new RPTCustomerReminderEntity();
        }
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        int orderNewQty = entity.getOrderNewQty();
        int reminderQty = entity.getReminderQty();
        int reminderFirstQty = entity.getReminderFirstQty();
        int reminderMultipleQty = entity.getReminderMultipleQty();
        int exceed48hourReminderQty = entity.getExceed48hourReminderQty();
        if (orderNewQty != 0) {
            entity.setReminderRate(numberFormat.format((float) reminderQty / orderNewQty * 100));

            entity.setReminderFirstRate(numberFormat.format((float) reminderFirstQty / orderNewQty * 100));

            entity.setReminderMultipleRate(numberFormat.format((float) reminderMultipleQty / orderNewQty * 100));

            entity.setExceed48hourReminderRate(numberFormat.format((float) exceed48hourReminderQty / orderNewQty * 100));
        }

        int reminderOrderQty = entity.getReminderOrderQty();
        if (reminderOrderQty != 0) {
            entity.setComplete24hourRate(numberFormat.format((float) entity.getComplete24hourQty() / reminderOrderQty * 100));

            entity.setOver48ReminderCompletedRate(numberFormat.format((float) entity.getOver48ReminderCompletedQty() / reminderOrderQty * 100));
        }

        return entity;
    }
}
