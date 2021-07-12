package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CustomerRevenueRptService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
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
public class CustomerRevenueRptTasks {

    @Autowired
    private CustomerRevenueRptService customerRevenueRptService;

    @Scheduled(cron = "${rpt.jobs.writeCustomerRevenueRptJob}")
    public void autoWriteCustomerRevenueRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }

        //写入前一天数据
        Date date = DateUtils.addDays(new Date(), -1);
        customerRevenueRptService.saveCustomerRevenueToRptDB(date);
    }
}
