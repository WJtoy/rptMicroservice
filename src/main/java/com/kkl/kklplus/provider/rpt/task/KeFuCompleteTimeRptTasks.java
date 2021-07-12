package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.KeFuCompleteTimeRptService;
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
public class KeFuCompleteTimeRptTasks {
    @Autowired
    private KeFuCompleteTimeRptService keFuCompleteTimeRptService;


    @Scheduled(cron = "${rpt.jobs.writeKeFuCompleteTimeRptJob}")
    public void autoWriteKeFuCompleteTimeRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try{
            keFuCompleteTimeRptService.writeCreateOrderFromYesterday(date);
        }catch (Exception e){
            log.error("KeFuCompleteTimeRptTasks.writeCreateOrderFromYesterday:{}", Exceptions.getStackTraceAsString(e));
        }
        try{
            keFuCompleteTimeRptService.updateOrderCloseData(date);
        }catch (Exception e){
            log.error("KeFuCompleteTimeRptTasks.updateOrderCloseData:{}", Exceptions.getStackTraceAsString(e));
        }
        try{
            keFuCompleteTimeRptService.updateOrderCancelledData(date);
        }catch (Exception e){
            log.error("KeFuCompleteTimeRptTasks.updateOrderCancelledData:{}", Exceptions.getStackTraceAsString(e));
        }
        try{
            keFuCompleteTimeRptService.updateRptPlanType(date);
        }catch (Exception e){
            log.error("KeFuCompleteTimeRptTasks.updateRptPlanType:{}", Exceptions.getStackTraceAsString(e));
        }
        try{
            keFuCompleteTimeRptService.updateRptComplain(date);
        }catch (Exception e){
            log.error("KeFuCompleteTimeRptTasks.updateRptComplain:{}", Exceptions.getStackTraceAsString(e));
        }
    }
}
