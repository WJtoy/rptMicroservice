package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CustomerReminderRptService;
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
public class CustomerReminderRptTasks {
    @Autowired
    private CustomerReminderRptService customerReminderRptService;

    @Scheduled(cron = "${rpt.jobs.writeCustomerReminderRptJob}")
    public void writeCustomerReminderRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        try {
            customerReminderRptService.saveCustomerReminderToRptDB(DateUtils.addDays(today, -1));
        } catch (Exception e) {
            log.error("CustomerReminderRptTasks.saveCustomerReminderToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }

    }
}
