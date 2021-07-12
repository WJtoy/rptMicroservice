package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CrushAreaRptService;
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
public class CrushAreaRptTasks {

    @Autowired
    private CrushAreaRptService crushAreaRptService;

    @Scheduled(cron = "${rpt.jobs.writeCrushAreaRptJob}")
    public void writeCrushAreaRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try{
            crushAreaRptService.insertCrushAreaRpt(date);
        }catch (Exception e){
            log.error("CrushAreaRptTasks.writeCrushAreaRptJob:{}", Exceptions.getStackTraceAsString(e));
        }


    }
}
