package com.kkl.kklplus.provider.rpt.chart.task;

import com.kkl.kklplus.provider.rpt.chart.service.ServicePointQtyStatisticsService;
import com.kkl.kklplus.provider.rpt.chart.service.ServicePointStreetQtyService;
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
public class ServicePointQtyChartTasks {

    @Autowired
    private ServicePointQtyStatisticsService servicePointQtyStatisticsService;

    @Autowired
    private ServicePointStreetQtyService servicePointStreetQtyService;
    @Scheduled(cron = "${rpt.jobs.writeServicePointQtyChartJob}")
    public void autoWriteServicePointQtyChartJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        //写入前一天数据
        Date date = DateUtils.addDays(new Date(), -1);
        servicePointQtyStatisticsService.saveServicePointQtyToRptDB(date);

        servicePointStreetQtyService.saveServicePointStreetQtyToRptDB(date);
    }
}
