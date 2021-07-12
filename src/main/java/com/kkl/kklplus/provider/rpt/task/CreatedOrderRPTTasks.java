package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.CreatedOrderService;
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
public class CreatedOrderRPTTasks {

    @Autowired
    private CreatedOrderService createdOrderService;

    @Scheduled(cron = "${rpt.jobs.writeCreatedOrderRptJob}")
    public void writeCreatedOrderRptJob(){
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date date = DateUtils.addDays(new Date(), -1);
        try {
            createdOrderService.deleteRepeatOrder(date);
            log.info("定时任务：deleteRepeatOrder删除重复数据成功");
        }catch (Exception e){
            log.error("定时任务：deleteRepeatOrder错误:{}","定时任务中间表删除重复数据错误",e);
        }
        try {
            createdOrderService.replenishOrder(date);
            log.info("定时任务：replenishOrder写入成功");
        }catch (Exception e){
            log.error("定时任务：replenishOrder错误:{}","定时任务中间表写入错误",e);
        }
    }
}
