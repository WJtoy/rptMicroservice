package com.kkl.kklplus.provider.rpt.task;

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
public class GradeQtyDailyRptTasks {

    @Autowired
    private GradeQtyDailyRptService gradeQtyDailyRptService;

    @Scheduled(cron = "${rpt.jobs.gradeQtyDailyRptJob}")
    public void gradeQtyDailyRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);

        try {
            gradeQtyDailyRptService.writeYesterdayQty(date);
        }catch (Exception e){
            log.error("CustomerReminderRptTasks.saveCustomerReminderToRptDB:{}", Exceptions.getStackTraceAsString(e));

        }

    }


}
