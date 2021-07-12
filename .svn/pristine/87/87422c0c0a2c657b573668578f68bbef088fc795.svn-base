package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.AbnormalFinancialReviewRptService;
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
public class AbnormalFinancialRptTasks {

    @Autowired
    private AbnormalFinancialReviewRptService abnormalFinancialReviewRptService;

    @Scheduled(cron = "${rpt.jobs.writeAbnormalFinancialRptJob}")
    public void writeAbnormalFinancialRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        try {
            abnormalFinancialReviewRptService.saveAbnormalFinancialReviewRptDB(DateUtils.addDays(today, -1));
        } catch (Exception e) {
            log.error("AbnormalFinancialRptTasks.writeAbnormalFinancialRptJob:{}", Exceptions.getStackTraceAsString(e));
        }

    }
}
