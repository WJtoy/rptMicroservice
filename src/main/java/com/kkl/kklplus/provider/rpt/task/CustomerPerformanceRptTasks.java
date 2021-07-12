package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CustomerPerformanceRptService;
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
public class CustomerPerformanceRptTasks {

    @Autowired
    CustomerPerformanceRptService customerPerformanceRptService;

    @Scheduled(cron = "${rpt.jobs.writeCustomerPerformanceRptJob}")
    public void writeCustomerPerformanceRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }

        Date lastMonth = DateUtils.addMonth(new Date(), -1);
            try {
                customerPerformanceRptService.writeYesterdayQty(DateUtils.getYear(lastMonth), DateUtils.getMonth(lastMonth));
            } catch (Exception e) {
                log.error("CustomerPerformanceRptTasks.writeCustomerPerformanceRptJob:{}", Exceptions.getStackTraceAsString(e));
            }
    }

    @Scheduled(cron = "${rpt.jobs.writeCustomerPerformanceCurrentMonthRptJob}")
    public void writeCustomerPerformanceMonthRptJob() {
        Date date = new Date();
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }

        try {
            customerPerformanceRptService.writeYesterdayQty(DateUtils.getYear(date), DateUtils.getMonth(date));
        } catch (Exception e) {
            log.error("CustomerPerformanceRptTasks.writeCustomerPerformanceCurrentMonthRptJob:{}", Exceptions.getStackTraceAsString(e));
        }
    }

}
