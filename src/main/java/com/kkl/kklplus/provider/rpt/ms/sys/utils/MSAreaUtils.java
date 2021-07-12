package com.kkl.kklplus.provider.rpt.ms.sys.utils;

import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.utils.SpringContextHolder;

import java.util.List;
import java.util.Map;

/**
 * 区域工具类
 */
public class MSAreaUtils {

    private static MSAreaService msAreaService = SpringContextHolder.getBean(MSAreaService.class);


    /**
     * 按区域类型返回所有区域信息
     */
    public static Map<Long, RPTArea> getAreaMap(int type) {
        List<RPTArea> areaList = msAreaService.getAreaByType(type);
        Map<Long, RPTArea> areaMap = Maps.newHashMap();
        for (RPTArea item : areaList) {
            areaMap.put(item.getId(), item);
        }
        return areaMap;
    }

    /**
     * 返回省份列表
     */
    public static List<RPTArea> getProvinceList() {
        return msAreaService.getAreaByType(RPTArea.TYPE_VALUE_PROVINCE);
    }

    /**
     * 查找突击覆盖的四级区域列表
     */
    public static List<Long> getCrushAreaMap(List<Long> productCategoryIds) {
        List<Long> areaList = msAreaService.findCoverAreaList(productCategoryIds);
        return areaList;
    }

    /**
     * 查找远程覆盖的四级区域列表
     */
    public static List<Long> getTravelAreaMap(List<Long> productCategoryIds) {
        List<Long> areaList = msAreaService.findTravelAreaList(productCategoryIds);
        return areaList;
    }

}
