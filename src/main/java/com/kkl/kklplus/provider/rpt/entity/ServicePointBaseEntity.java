package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;

import lombok.Getter;
import lombok.Setter;

import java.util.Date;

public class ServicePointBaseEntity extends RPTBase {


    @Getter
    @Setter
    private Long id;

    @Getter
    @Setter
    private String engineerName = " ";

    @Getter
    @Setter
    private String servicePointNo = "";   //网点编号

    @Getter
    @Setter
    private String name = "";   //网点名称

    @Getter
    @Setter
    private String contactInfo1 = ""; //联系方式(手机)

    @Getter
    @Setter
    private String contactInfo2 = ""; //联系方式(电话)

    @Getter
    @Setter
    private String contactInfo = ""; //安维人员联系方式(电话)

    @Getter
    @Setter
    private Date contractDate; //签约日期

    @Getter
    @Setter
    private Integer delFlag=0;

    @Getter
    @Setter
    private Integer discountFlag = 0;

    @Getter
    @Setter
    private Integer autoPlanFlag = 0; //自动派单开关,0-人工派单,1:自动派单 //2019-4-9

    @Getter
    @Setter
    private Integer insuranceFlag = 0;//是否开启保险计算Flag;0 不开启 1 开启

    @Getter
    @Setter
    private Integer timeLinessFlag = 0;//网点时效开关//2018-06-25  时效开关默认为开 //2018-7-31 时效开关默认为关闭

    @Getter
    @Setter
    private int useDefaultPrice = 0; //使用默认价，0:自定义价, >0:数据字典维护价格

    @Getter
    @Setter
    private Long primaryId;

    @Getter
    @Setter
    private String address = "";

    @Getter
    @Setter
    private Integer level;  //等级

    @Getter
    @Setter
    private Integer signFlag = -1;  //是否签约

    @Getter
    @Setter
    private String remarks = "";	// 备注

    @Getter
    @Setter
    private int orderCount = 0;

    @Getter
    @Setter
    private int grade = 0; //评价分数

    @Getter
    @Setter
    private Integer status;

    @Getter
    @Setter
    private Long areaId;

    @Getter
    @Setter
    private Integer bank;

    @Getter
    @Setter
    private String bankNo;

    @Getter
    @Setter
    private String bankOwner;

    @Getter
    @Setter
    private Integer paymentType;



}
