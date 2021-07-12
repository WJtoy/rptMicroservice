package com.kkl.kklplus.provider.rpt.utils.web;

import com.google.common.collect.Lists;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.reflect.TypeToken;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

public class RPTOrderItemUtils {

    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTOrderItem.class, RPTOrderItemAdapter.getInstance()).create();

    /**
     * OrderItem列表转成json字符串
     */
    public static String toOrderItemsJson(List<RPTOrderItem> orderItems) {
        String json = null;
        if (orderItems != null && orderItems.size() > 0) {
            json = gson.toJson(orderItems, new TypeToken<List<RPTOrderItem>>() {
            }.getType());
            /**
             *  因为myCat1.6不支持在json或text类型的字段中存储英文括号，故将所有的英文括号替换成中文括号.
             */
            json = json.replace("(", "（");
            json = json.replace(")", "）");
        }
        return json;
    }

    /**
     * json字符串转成OrderItem列表
     */
    public static List<RPTOrderItem> fromOrderItemsJson(String json) {
        List<RPTOrderItem> orderItems = null;
        if (StringUtils.isNotEmpty(json)) {
            orderItems = gson.fromJson(json, new TypeToken<List<RPTOrderItem>>() {
            }.getType());
        }
        return orderItems != null ? orderItems : Lists.newArrayList();
    }

    /**
     * 设置列表中orderitem的某些属性，如服务类型、产品
     *
     * @param orderItems     OrderItem列表
     * @param cacheDataTypes CacheDataTypeEnum的集合
     * @return 设置好属性的OrderItem列表
     */
    public static List<RPTOrderItem> setOrderItemProperties(List<RPTOrderItem> orderItems, Set<CacheDataTypeEnum> cacheDataTypes) {
        if (orderItems != null && orderItems.size() > 0 && cacheDataTypes != null && cacheDataTypes.size() > 0) {
            Map<Long, RPTServiceType> serviceTypeMap = MDUtils.getAllServiceTypeMap();
            Set<Long> productIds = orderItems.stream().map(i -> i.getProduct().getId()).collect(Collectors.toSet());
            Map<Long, RPTProduct> productMap = MDUtils.getAllProductMap(Lists.newArrayList(productIds));
            RPTServiceType serviceType;
            RPTProduct product;
            for (RPTOrderItem item : orderItems) {
                if (CacheDataTypeEnum.isExists(cacheDataTypes, CacheDataTypeEnum.SERVICETYPE)) {
                    serviceType = serviceTypeMap.get(item.getServiceType().getId());
                    item.setServiceType(serviceType);
                }
                if (CacheDataTypeEnum.isExists(cacheDataTypes, CacheDataTypeEnum.PRODUCT)) {
                    product = productMap.get(item.getProductId());
                    item.setProduct(product);
                }
            }
        }
        return orderItems;
    }

}
