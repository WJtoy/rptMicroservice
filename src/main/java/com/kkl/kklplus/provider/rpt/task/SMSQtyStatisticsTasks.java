package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.SMSQtyStatisticsService;
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
public class SMSQtyStatisticsTasks {
    @Autowired
    private SMSQtyStatisticsService smsQtyStatisticsService;

    @Scheduled(cron = "${rpt.jobs.writeYesterdayMessageRptJob}")
    public void autoWriteYesterdayMessageRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        //检查前一天的前一天数据
        smsQtyStatisticsService.checkEveShortMessageQty();

        //写入前一天数据
        Date date = DateUtils.addDays(new Date(), -1);
        smsQtyStatisticsService.writeYesterdayMessage(date);
    }
}
