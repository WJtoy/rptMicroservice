package com.kkl.kklplus.provider.rpt.entity;

import java.util.Set;
import java.util.stream.Collectors;

public enum CacheDataTypeEnum {

    /**
     * 缓存类型
     */
    CUSTOMER(1, "客户基本资料"),
    SERVICEPOINT(2, "网点基本资料"),
    ENGINEER(32, "师傅基础资料"),

    SERVICETYPE(4, "服务类型"),
    PRODUCT(8, "产品"),

    EXPPRESSCOMPANY(16, "快递公司");

    public long value;
    public String name;

    private CacheDataTypeEnum(long value, String name) {
        this.value = value;
        this.name = name;
    }

    public static boolean isExists(Set<CacheDataTypeEnum> cacheDataTypes, CacheDataTypeEnum cacheDataType) {
        Set<CacheDataTypeEnum> cacheDataTypeEnums = cacheDataTypes.stream().filter(i -> i.value == cacheDataType.value).collect(Collectors.toSet());
        return cacheDataTypeEnums.size() > 0;
    }

}
