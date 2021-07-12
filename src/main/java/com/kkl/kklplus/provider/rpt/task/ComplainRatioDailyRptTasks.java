package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.ComplainRatioDailyRptService;
import com.kkl.kklplus.provider.rpt.service.GradeQtyDailyRptService;
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
public class ComplainRatioDailyRptTasks {

    @Autowired
    private ComplainRatioDailyRptService complainRatioDailyRptService;

    @Scheduled(cron = "${rpt.jobs.complainRatioDailyRptJob}")
    public void complainRatioDailyRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = new Date();

        try {
            complainRatioDailyRptService.updateComplainOrderToRptDB(DateUtils.addDays(date, -2));
        } catch (Exception e) {
            log.error("ComplainRatioDailyRptTasks.updateComplainOrderToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            complainRatioDailyRptService.saveComplainOrderRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("ComplainRatioDailyRptTasks.saveComplainOrderRptDB:{}", Exceptions.getStackTraceAsString(e));
        }

    }

}
