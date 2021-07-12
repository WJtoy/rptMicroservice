package com.kkl.kklplus.provider.rpt.common.service;

import com.github.benmanes.caffeine.cache.Cache;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.mapper.AreaCacheMapper;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class AreaCacheService {

    private static final String CACHE_KEY_ALL_STREET_MAP = "ALL:STREET:MAP";

    private static final String CACHE_KEY_ALL_COUNTY_MAP = "ALL:COUNTY:MAP";

    private static final String CACHE_KEY_ALL_CITY_MAP = "ALL:CITY:MAP";

    private static final String CACHE_KEY_ALL_PROVINCE_MAP = "ALL:PROVINCE:MAP";

    @Resource
    private AreaCacheMapper areaCacheMapper;

    @Autowired
    private Cache<String, Object> caffeineCache;

    /**
     * 获取系统所有的乡镇
     */
    public Map<Long,RPTArea> getAllTownMap() {
        Map<Long,RPTArea> allTownMap = (Map<Long,RPTArea>)caffeineCache.asMap().get(CACHE_KEY_ALL_STREET_MAP);
        if (allTownMap == null) {
            List<RPTArea> list = areaCacheMapper.getAllTownList();
            allTownMap = list.stream().collect(Collectors.toMap(RPTBase::getId, i->i));
            caffeineCache.put(CACHE_KEY_ALL_STREET_MAP, allTownMap);
        }
        return allTownMap;
    }

    public Map<Long,RPTArea> getAllCountyMap() {
        Map<Long,RPTArea> allCountyMap = (Map<Long,RPTArea>)caffeineCache.asMap().get(CACHE_KEY_ALL_COUNTY_MAP);
        if (allCountyMap == null) {
            List<RPTArea> list = areaCacheMapper.getAllCountyList();
            allCountyMap = list.stream().collect(Collectors.toMap(RPTBase::getId, i->i));
            caffeineCache.put(CACHE_KEY_ALL_COUNTY_MAP, allCountyMap);
        }
        return allCountyMap;
    }



    public Map<Long,RPTArea> getAllCityMap() {
        Map<Long,RPTArea> allCityMap = (Map<Long,RPTArea>)caffeineCache.asMap().get(CACHE_KEY_ALL_CITY_MAP);
        if (allCityMap == null) {
            List<RPTArea> list = areaCacheMapper.getAllCityList();
            allCityMap = list.stream().collect(Collectors.toMap(RPTBase::getId, i->i));
            caffeineCache.put(CACHE_KEY_ALL_CITY_MAP, allCityMap);
        }
        return allCityMap;
    }



    public Map<Long,RPTArea> getAllProvinceMap() {
        Map<Long,RPTArea> allProvinceMap = (Map<Long,RPTArea>)caffeineCache.asMap().get(CACHE_KEY_ALL_PROVINCE_MAP);
        if (allProvinceMap == null) {
            List<RPTArea> list = areaCacheMapper.getAllProvinceList();
            allProvinceMap = list.stream().collect(Collectors.toMap(RPTBase::getId, i->i));
            caffeineCache.put(CACHE_KEY_ALL_PROVINCE_MAP, allProvinceMap);
        }
        return allProvinceMap;
    }


}
