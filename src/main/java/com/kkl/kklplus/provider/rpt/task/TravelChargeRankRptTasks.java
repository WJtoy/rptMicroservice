package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.TravelChargeRankRptService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Lazy;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.util.Date;

@Component
@Slf4j
@Lazy(value = false)
public class TravelChargeRankRptTasks {

    @Autowired
    private TravelChargeRankRptService travelChargeRankRptService;

    //@Scheduled(cron = "0 0 4 * * ?")
    @Scheduled(cron = "${rpt.jobs.writTravelChargeRankRptJob}")
    public void autoWriteTravelChargeRankRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try {
            travelChargeRankRptService.insertTravelChargeRank(date);
        }catch (Exception e){
            log.error("TravelChargeRankRptTasks.autoWriteTravelChargeRankRptJob:{}", Exceptions.getStackTraceAsString(e));
        }
    }
}
