package com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils;

import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.entity.b2bcenter.md.B2BDataSourceEnum;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.service.B2BCustomerMappingService;
import com.kkl.kklplus.provider.rpt.utils.SpringContextHolder;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import org.javatuples.Pair;

import java.util.List;
import java.util.Map;

public class B2BCenterUtils {

    private static B2BCustomerMappingService customerMappingService = SpringContextHolder.getBean(B2BCustomerMappingService.class);


    public static String getShopName(Long customerId, String shopId) {
        String shopName = "";
        if (customerId != null && customerId > 0 && StringUtils.isNotBlank(shopId)) {
            shopName = customerMappingService.getshopName(customerId, shopId);
        }
        return shopName;
    }


    /**
     * 返回元祖对象，包含两个Map：
     * 第一个Map的key为shopId，value为B2BCustomerMapping；
     * 第二个Map的key为dataSourceId:shopId，value为B2BCustomerMapping
     */
    public static Pair<Map<String, B2BCustomerMapping>, Map<String, B2BCustomerMapping>> getAllCustomerMappingMaps() {
        Map<String, B2BCustomerMapping> resultA = Maps.newHashMap();
        Map<String, B2BCustomerMapping> resultB = Maps.newHashMap();
        List<B2BCustomerMapping> list = customerMappingService.getAllCustomerMapping();
        if (list.size() > 0) {
            for (B2BCustomerMapping item : list) {
                if (item != null && StringUtils.isNotBlank(item.getShopId())) {
                    resultA.put(item.getShopId(), item);
                    if (B2BDataSourceEnum.isB2BDataSource(item.getDataSource())) {
                        resultB.put(String.format("%d:%s", item.getDataSource(), item.getShopId()), item);
                    }
                }
            }
        }
        return new Pair<>(resultA, resultB);
    }

    /**
     * 返回元祖对象，包含两个Map：
     * 第一个Map的key为shopId，value为ShopName；（dataSource == 1时，使用该对象）
     * 第二个Map的key为dataSourceId:shopId，value为ShopName （dataSource ==  2时，使用该对象）
     */
    public static Pair<Map<String, String>, Map<String, String>> getAllShopMaps() {
        Map<String, String> resultA = Maps.newHashMap();
        Map<String, String> resultB = Maps.newHashMap();
        List<B2BCustomerMapping> list = customerMappingService.getAllCustomerMapping();
        if (list.size() > 0) {
            for (B2BCustomerMapping item : list) {
                if (item != null && StringUtils.isNotBlank(item.getShopId())) {
                    resultA.put(item.getShopId(), item.getShopName());
                    if (B2BDataSourceEnum.isB2BDataSource(item.getDataSource())) {
                        resultB.put(String.format("%d:%s", item.getDataSource(), item.getShopId()), item.getShopName());
                    }
                }
            }
        }
        return new Pair<>(resultA, resultB);
    }

    /**
     * 获取店铺名称
     *
     * @param allShopMap 元祖对象，包含两个Map：第一个Map的key为shopId，value为ShopName；（dataSource == 1时，使用该对象）;第二个Map的key为dataSourceId:shopId，value为ShopName （dataSource ==  2时，使用该对象）
     */
    public static String getShopName(Integer dataSourceId, String shopId, Pair<Map<String, String>, Map<String, String>> allShopMap) {
        String shopName = "";
        if (B2BDataSourceEnum.isDataSource(dataSourceId) && StringUtils.isNotBlank(shopId)) {
            if (B2BDataSourceEnum.isB2BDataSource(dataSourceId)) {
                shopName = allShopMap.getValue1().get(String.format("%s:%s", dataSourceId, shopId));
            } else {
                shopName = allShopMap.getValue0().get(shopId);
            }
        }
        return StringUtils.toString(shopName);
    }


    /**
     * 查询所有的数据源中的店铺与客户的对应关系
     *
     * @return key为“shopId”
     */
    public static Map<String, B2BCustomerMapping> getAllCustomerMappingMapForKKL() {
        Map<String, B2BCustomerMapping> result = Maps.newHashMap();
        List<B2BCustomerMapping> list = customerMappingService.getAllCustomerMapping();
        if (list.size() > 0) {
            for (B2BCustomerMapping item : list) {
                if (item != null && StringUtils.isNotBlank(item.getShopId())) {
                    result.put(item.getShopId(), item);
                }
            }
        }
        return result;
    }

    /**
     * 查询所有的数据源中的店铺与客户的对应关系
     *
     * @return key为“dataSourceId:shopId”
     */
    public static Map<String, B2BCustomerMapping> getAllCustomerMappingMapForB2B() {
        Map<String, B2BCustomerMapping> result = Maps.newHashMap();
        List<B2BCustomerMapping> list = customerMappingService.getAllCustomerMapping();
        if (list.size() > 0) {
            for (B2BCustomerMapping item : list) {
                if (item != null && B2BDataSourceEnum.isB2BDataSource(item.getDataSource()) && StringUtils.isNotBlank(item.getShopId())) {
                    result.put(String.format("%d:%s", item.getDataSource(), item.getShopId()), item);
                }
            }
        }
        return result;
    }


}
