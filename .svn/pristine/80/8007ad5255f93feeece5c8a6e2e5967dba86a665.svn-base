package com.kkl.kklplus.provider.rpt.config;

import lombok.Getter;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * @author Zhoucy
 * @date 2018/7/31 14:34
 **/
@ConfigurationProperties(prefix = "rpt")
public class ProviderRptProperties {

    /**
     * 系统标识
     */
    @Getter
    @Setter
    private String systemCode = "";

    /**
     * 是否开启定时任务
     */
    @Getter
    @Setter
    private Boolean scheduleEnabled = false;

    /**
     * 是否开启更新平台服务费
     */
    @Getter
    @Setter
    private Boolean infoFeeEnabled = false;

    /**
     * Web系统的上线时间
     */
    @Getter
    @Setter
    private String goLiveDate = "";

    /**
     * Web数据库的分片数量
     */
    @Getter
    @Setter
    private Integer quarterQty = 0;

    /**
     * 报表数据库的分片数量
     */
    @Getter
    @Setter
    private Integer rptQuarterQty = 0;

    /**
     * 平台服务费
     */
    @Getter
    @Setter
    private double PlatformFeeRate = 0;
    /**
     * 报表微服务的上线时间
     */
    @Getter
    @Setter
    private String rptGoLiveDate = "";

    @Getter
    private final RptExcelFileDir rptExcelFileDir = new RptExcelFileDir();

    public static class RptExcelFileDir {
        @Getter
        @Setter
        private String host = "";
        @Getter
        @Setter
        private String uploadDir = "";
    }

}
