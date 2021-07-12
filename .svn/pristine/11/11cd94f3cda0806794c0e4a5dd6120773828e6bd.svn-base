package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CustomerOrderTimeRptService;
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
public class CustomerOrderTimeRptTasks {
    @Autowired
    private CustomerOrderTimeRptService customerOrderTimeRptService;


    @Scheduled(cron = "${rpt.jobs.writeCustomerOrderTimeRptJob}")
    public void autoWriteCustomerOrderTimeRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try{
            customerOrderTimeRptService.saveCustomerOrderTimeToRptDB(date);
        }catch (Exception e){
            log.error("CustomerOrderTimeRptTasks.saveCustomerOrderTimeToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }
}
