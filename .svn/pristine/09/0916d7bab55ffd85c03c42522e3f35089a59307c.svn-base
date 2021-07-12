package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.GradedOrderRptService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Lazy;
import org.springframework.context.annotation.PropertySource;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.util.Date;

@Component
@Slf4j
@Lazy(value = false)
@PropertySource(value = "classpath:application.yml")
public class GradedOrderRptTasks {

    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    @Scheduled(cron = "${rpt.jobs.writeGradedOrderRptJob}")
    public void writeGradedOrderRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        try {
            gradedOrderRptService.deleteHavingGradedOrder(DateUtils.addDays(today,-1));
        }catch (Exception e){
            log.error("GradedOrderRptTasks.deleteHavingGradedOrder:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            gradedOrderRptService.saveMissGradedOrdersToRptDB(DateUtils.addDays(today, -1));
        } catch (Exception e) {
            log.error("GradedOrderRptTasks.saveMissGradedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
//        try {
//            gradedOrderRptService.saveGradedOrderToRptDB(DateUtils.addDays(today, -1));
//        } catch (Exception e) {
//            log.error("GradedOrderRptTasks.saveGradedOrderToRptDB:{}", Exceptions.getStackTraceAsString(e));
//        }
    }

}
