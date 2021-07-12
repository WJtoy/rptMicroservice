package com.kkl.kklplus.provider.rpt.utils;

import com.kkl.kklplus.entity.rpt.common.RPTSystemCodeEnum;
import com.kkl.kklplus.provider.rpt.config.ProviderRptProperties;

import java.util.Date;

/**
 * @author Zhoucy
 * @date 2018/8/27 9:59
 **/
public class RptCommonUtils {

    private static ProviderRptProperties providerRptProperties = SpringContextHolder.getBean(ProviderRptProperties.class);

    /**
     * 获取当前系统标识
     */
    public static int getSystemId() {
        String systemCode = providerRptProperties.getSystemCode();
        return RPTSystemCodeEnum.get(systemCode).id;
    }

    public static int getQuarterQty() {
        Integer quarterQty = providerRptProperties.getQuarterQty();
        return quarterQty == null ? 0 : quarterQty;
    }

    public static boolean scheduleEnabled() {
        return providerRptProperties.getScheduleEnabled();
    }

    public static Date getGoLiveDate() {
        Date goLiveDate = null;
        String date = providerRptProperties.getGoLiveDate();
        if (StringUtils.isNotBlank(date)) {
            goLiveDate = DateUtils.parseDate(date);
        }
        if (goLiveDate == null) {
            goLiveDate = DateUtils.getDate(2019, 1, 1);
        }
        return goLiveDate;
    }

    public static int getRptQuarterQty() {
        Integer quarterQty = providerRptProperties.getRptQuarterQty();
        return quarterQty == null ? 0 : quarterQty;
    }

    public static double getPlatformFeeRate() {
        return providerRptProperties.getPlatformFeeRate();
    }

    public static boolean getInfoFeeEnabled() {
        return providerRptProperties.getInfoFeeEnabled();
    }

    public static Date getRptGoLiveDate() {
        Date goLiveDate = null;
        String date = providerRptProperties.getRptGoLiveDate();
        if (StringUtils.isNotBlank(date)) {
            goLiveDate = DateUtils.parseDate(date);
        }
        if (goLiveDate == null) {
            goLiveDate = DateUtils.getDate(2017, 1, 1);
        }
        return goLiveDate;
    }
}
