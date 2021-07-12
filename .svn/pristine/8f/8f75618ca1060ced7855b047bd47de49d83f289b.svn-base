package com.kkl.kklplus.provider.rpt.task;

import com.kkl.kklplus.provider.rpt.service.*;
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

/**
 * 完工明细、退补单、客户退补、取消/退单
 */
@Component
@Slf4j
@Lazy(value = false)
@PropertySource(value = "classpath:application.yml")
public class CustomerChargeRptTasks {

    @Autowired
    private CustomerChargeSummaryRptService customerChargeSummaryRptService;
    @Autowired
    private CompletedOrderRptService completedOrderRptService;
    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;
    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;
    @Autowired
    private ServicePointCompletedOrderRptService servicePointCompletedOrderRptService;
    @Autowired
    private ServicePointWriteOffRptService servicePointWriteOffRptService;

    @Autowired
    private CustomerChargeSummaryRptNewService customerChargeSummaryRptNewService;

    //    @Scheduled(cron ="0 0 3 1 * ?")
//    @Scheduled(cron = "${rpt.jobs.writeCustomerChargeSummaryRptJob}")
//    public void writeCustomerChargeSummaryRptJob() {
//        if (!RptCommonUtils.scheduleEnabled()) {
//            return;
//        }
//        Date today = new Date();
//        saveCustomerChargeSummary(today);
//    }

    //    @Scheduled(cron = "0 0 3 2,3 * ?")
    @Scheduled(cron = "${rpt.jobs.rewriteCustomerChargeSummaryRptJob}")
    public void rewriteCustomerChargeSummaryRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        updateCustomerChargeSummary(today);
    }

    @Scheduled(cron = "${rpt.jobs.rewriteCustomerChargeSummaryCurrentMonthRptJob}")
    public void rewriteCustomerChargeSummaryCurrentMonthRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        updateCustomerChargeSummaryCurrentMonth(today);
    }

    //    @Scheduled(cron = "0 30 3 * * ?")
    @Scheduled(cron = "${rpt.jobs.writeCustomerChargeRptJob}")
    public void writeCustomerChargeRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date today = new Date();
        saveCompletedOrders(today);
        saveCancelledOrders(today);
        saveCustomerWriteOffs(today);
        saveServicePointCompletedOrders(today);
        saveServicePointWriteOffs(today);
    }

//    private void saveCustomerChargeSummary(Date date) {
//        Date lastMonth = DateUtils.addMonth(date, -1);
//        try {
//            customerChargeSummaryRptService.saveCustomerChargeSummaryToRptDB(DateUtils.getYear(lastMonth), DateUtils.getMonth(lastMonth));
//        } catch (Exception e) {
//            log.error("CustomerChargeRptTasks.saveCustomerChargeSummaryToRptDB:{}", Exceptions.getStackTraceAsString(e));
//        }
//    }

    private void updateCustomerChargeSummary(Date date) {
        Date lastMonth = DateUtils.addMonth(date, -1);
        try {
            customerChargeSummaryRptNewService.updateCustomerChargeSummaryToRptDB(DateUtils.getYear(lastMonth), DateUtils.getMonth(lastMonth));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.updateCustomerChargeSummaryToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    private void updateCustomerChargeSummaryCurrentMonth(Date date) {
        try {
            customerChargeSummaryRptNewService.updateCustomerChargeSummaryToRptDB(DateUtils.getYear(date), DateUtils.getMonth(date));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.updateCustomerChargeSummaryCurrentMonth:{}", Exceptions.getStackTraceAsString(e));
        }
    }



    private void saveCompletedOrders(Date date) {
        try {
            completedOrderRptService.saveMissedCompletedOrdersToRptDB(DateUtils.addDays(date, -2));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveMissedCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            completedOrderRptService.saveCompletedOrdersToRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    private void saveCancelledOrders(Date date) {
//        try {
//            cancelledOrderRptService.saveMissedCancelledOrdersToRptDB(DateUtils.addDays(date, -2));
//        } catch (Exception e) {
//            log.error("CustomerChargeRptTasks.saveMissedCancelledOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
//        }
        try {
            cancelledOrderRptService.saveMissedCancelledOrdersToRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveCancelledOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    private void saveCustomerWriteOffs(Date date) {
        try {
            customerWriteOffRptService.saveMissedCustomerWriteOffsToRptDB(DateUtils.addDays(date, -2));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveMissedCustomerWriteOffsToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            customerWriteOffRptService.saveCustomerWriteOffToRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveCustomerWriteOffToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    private void saveServicePointCompletedOrders(Date date) {
        try {
            servicePointCompletedOrderRptService.saveMissedServicePointCompletedOrdersToRptDB(DateUtils.addDays(date, -2));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveMissedServicePointCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            servicePointCompletedOrderRptService.saveServicePointCompletedOrdersToRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveServicePointCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    private void saveServicePointWriteOffs(Date date) {
        try {
            servicePointWriteOffRptService.saveMissedServicePointWriteOffsToRptDB(DateUtils.addDays(date, -2));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveMissedServicePointWriteOffsToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
        try {
            servicePointWriteOffRptService.saveServicePointWriteOffToRptDB(DateUtils.addDays(date, -1));
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveServicePointWriteOffToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }


/*
    @Scheduled(cron = "${rpt.jobs.writeCustomerChargeSummaryRptJob}")
    public void writeCustomerChargeSummaryRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date beginDate = DateUtils.parseDate("2019-1-1");
        Date endDate = DateUtils.parseDate("2019-11-14");
        List<Date> dateList = Lists.newArrayList();
        while (beginDate.getTime() < endDate.getTime()) {
            dateList.add(beginDate);
            beginDate = DateUtils.addMonth(beginDate, 1);
        }
        for (Date date : dateList) {
            saveCustomerChargeSummary(date);
        }
    }

    @Scheduled(cron = "${rpt.jobs.rewriteCustomerChargeSummaryRptJob}")
    public void rewriteCustomerChargeSummaryRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }

        Date beginDate = DateUtils.parseDate("2019-1-1");
        Date endDate = DateUtils.parseDate("2019-11-14");
        List<Date> dateList = Lists.newArrayList();
        while (beginDate.getTime() < endDate.getTime()) {
            dateList.add(beginDate);
            beginDate = DateUtils.addMonth(beginDate, 1);
        }
        for (Date date : dateList) {
            updateCustomerChargeSummary(date);
        }
    }

    @Scheduled(cron = "${rpt.jobs.writeCustomerChargeRptJob}")
    public void writeCustomerChargeRptJob() {
        if (!RptCommonUtils.scheduleEnabled()) {
            return;
        }
        Date beginDate = DateUtils.parseDate("2019-2-1");
        Date endDate = DateUtils.parseDate("2019-11-14");
        List<Date> dateList = Lists.newArrayList();
        while (beginDate.getTime() < endDate.getTime()) {
            dateList.add(beginDate);
            beginDate = DateUtils.addDays(beginDate, 1);
        }
        for (Date date : dateList) {
            saveCompletedOrders(date);
            saveCancelledOrders(date);
            saveCustomerWriteOffs(date);
        }
    }

*/
}
