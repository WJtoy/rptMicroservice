package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CancelledOrderDailyRptService;
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
public class CancelledOrderDailyRptTasks {

    @Autowired
    private CancelledOrderDailyRptService cancelledOrderDailyRptService;

    @Scheduled(cron = "${rpt.jobs.writeCancelledOrderDailyRptJob}")
    public void writeCancelledOrderDailyRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try {
            cancelledOrderDailyRptService.deleteRepeatOrder(date);
            log.info("定时任务：cancelledOrderDailyRptService.deleteRepeatOrder删除每日退单取消单中间表重复数据成功");
        }catch (Exception e){
            log.error("定时任务：cancelledOrderDailyRptService.deleteRepeatOrder错误:{}","定时任务退单取消单中间表删除重复数据错误",e);
        }
        try {
            cancelledOrderDailyRptService.saveMissCancelledOrderRptToDB(date);
            log.info("定时任务：cancelledOrderDailyRptService.saveMissCancelledOrderRptToDB写入成功");
        }catch (Exception e){
            log.error("定时任务：cancelledOrderDailyRptService.saveMissCancelledOrderRptToDB错误:{}","定时任务退单取消单中间表写入错误",e);
        }
    }
}
