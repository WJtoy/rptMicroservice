/**
 * Copyright &copy; 2012-2016 <a href="https://github.com/thinkgem/jeesite">JeeSite</a> All rights reserved.
 */
package com.kkl.kklplus.provider.rpt.utils;

import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * 数据库分片工具
 */
@Slf4j
public class QuarterUtils {

    /**
     * 按季度分片
     */
    public static String getSeasonQuarter(Date date) {
        String year = DateUtils.formatDate(date, "yyyy");
        int season = DateUtils.getSeason(date);
        return String.format("%s%s", year, season);
    }

    /**
     * 按季度分片
     */
    public static String getSeasonQuarter(Long millis) {
        Date date = DateUtils.longToDate(millis);
        String year = DateUtils.formatDate(date, "yyyy");
        int season = DateUtils.getSeason(date);
        return String.format("%s%s", year, season);
    }

    /**
     * 根据订单号获得数据库分片
     */
    public static String getOrderQuarterFromNo(String orderNo) {
        if (StringUtils.isBlank(orderNo)) {
            return "";
        }
        if (orderNo.trim().length() == 14) {
            try {
                String strDate = orderNo.trim().substring(1, 9);
                Date date = DateUtils.parse(strDate, "yyyyMMdd");
                int quarter = DateUtils.getSeason(date);
                return String.format("%s%s", DateUtils.getYear(date), quarter);
            } catch (Exception e) {
                return "";
            }
        }
        return "";
    }


    public static List<String> getQuarters(Date startDate, Date endDate) {
        List<String> quarters = Lists.newArrayList();
        if (startDate == null) {
            return quarters;
        }
        if (endDate == null || DateUtils.isGreaterNow(endDate)) {
            endDate = new Date();
        }
        startDate = DateUtils.parseDate(DateUtils.formatDate(startDate, "yyyy-MM-dd"));
        endDate = DateUtils.parseDate(DateUtils.formatDate(endDate, "yyyy-MM-dd"));
        int startMonth, endMonth;
        while (true) {
            startMonth = startDate.getMonth();
            endMonth = endDate.getMonth();
            if (startDate.getTime() > endDate.getTime()
                    && startMonth != endMonth
                    && (startMonth / 3 + 1) != (endMonth / 3 + 1)
            ) {
                break;
            }
            quarters.add(getSeasonQuarter(startDate));
            startDate = DateUtils.addMonth(startDate, 3);
        }
        return quarters;
    }

    /**
     * 获取分片名字列表
     */
    public static List<String> getQuarters(int quarterCount) {
        List<String> quarterList = Lists.newArrayList();
        if (quarterCount > 0) {
            Date goLiveDate = RptCommonUtils.getGoLiveDate();
            Date startDate = getStartDateOfSeason(goLiveDate);
            Date endDate = DateUtils.addMonth(startDate, 3 * quarterCount);
            endDate = DateUtils.addDays(endDate, -1);
            quarterList = getQuarters(startDate, endDate);
//            log.warn("开始日期:{},结束日期:{},分片:{}", DateUtils.formatDate(startDate, "yyyy-MM-dd"), DateUtils.formatDate(endDate, "yyyy-MM-dd"), quarterList);
        }
        return quarterList;
    }

    /**
     * 获取报表数据库分片名字列表
     */
    public static List<String> getRptQuarters() {
        List<String> quarterList = Lists.newArrayList();
        int quarterCount =  RptCommonUtils.getRptQuarterQty();
        Date goLiveDate = RptCommonUtils.getRptGoLiveDate();
        if (quarterCount > 0) {
            Date startDate = getStartDateOfSeason(goLiveDate);
            Date endDate = DateUtils.addMonth(startDate, 3 * quarterCount);
            endDate = DateUtils.addDays(endDate, -1);
            quarterList = getQuarters(startDate, endDate);
        }
        return quarterList;
    }

    private static Date getStartDateOfSeason(Date date) {
        int season = DateUtils.getSeason(date);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        switch (season) {
            case 1:
                calendar.set(Calendar.MONTH, 0);
                calendar.set(Calendar.DAY_OF_MONTH, 1);
                break;
            case 2:
                calendar.set(Calendar.MONTH, 3);
                calendar.set(Calendar.DAY_OF_MONTH, 1);
                break;
            case 3:
                calendar.set(Calendar.MONTH, 6);
                calendar.set(Calendar.DAY_OF_MONTH, 1);
                break;
            case 4:
                calendar.set(Calendar.MONTH, 9);
                calendar.set(Calendar.DAY_OF_MONTH, 1);
                break;
        }
        return calendar.getTime();
    }

}
