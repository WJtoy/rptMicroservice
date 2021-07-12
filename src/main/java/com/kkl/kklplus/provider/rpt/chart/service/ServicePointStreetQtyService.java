package com.kkl.kklplus.provider.rpt.chart.service;


import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.provider.rpt.chart.entity.RPTServicePointStreetQtyEntity;
import com.kkl.kklplus.provider.rpt.chart.mapper.ServicePointStreetQtyMapper;
import com.kkl.kklplus.provider.rpt.chart.ms.md.service.MSServicePointStreetService;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointStreetQtyService {

    @Autowired
    private MSServicePointStreetService msServicePointStreetService;

    @Resource
    private ServicePointStreetQtyMapper servicePointStreetQtyMapper;

    @Autowired
    private MSAreaService msAreaService;

    @Autowired
    private AreaCacheService areaCacheService;

    public Map<String, Object> getServicePointStreetQtyData(RPTDataDrawingListSearch search) {
        Map<String, Object> map = new HashMap<>();
        int systemId = RptCommonUtils.getSystemId();
        Long endDate = DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime();
        Long startDate = DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime();
        Long productCategoryId = 0L;
        RPTServicePointStreetQtyEntity servicePointStreetQty = servicePointStreetQtyMapper.getServicePointStreetQty(systemId, startDate, endDate, productCategoryId);
        if (servicePointStreetQty == null) {
            return map;
        }
        RPTServicePointStreetQtyEntity servicePointStreetAutoPlanQty = servicePointStreetQtyMapper.getServicePointStreetAutoPlanQty(systemId, startDate, endDate, productCategoryId);
        Map<Long, RPTProductCategory> allProductCategoryMap = MDUtils.getAllProductCategoryMap();
        List<Long> productCategoryIds = Lists.newArrayList(allProductCategoryMap.keySet());
        List<RPTServicePointStreetQtyEntity> servicePointProductCategoryQty = servicePointStreetQtyMapper.getServicePointProductCategoryQty(systemId, startDate, endDate, productCategoryIds);

        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        String productCategoryName = "";

        String frequentQtyRate;
        String trialQtyRate;
        String withoutQtyRate;
        String autoPlanStreetRate;
        long servicePointStreetTotal = servicePointStreetQty.getServicePointStreet();

        for (RPTServicePointStreetQtyEntity productCategoryEntity : servicePointProductCategoryQty) {
            if (productCategoryEntity.getProductCategoryId() != 0) {
                if (!allProductCategoryMap.isEmpty()) {
                    if (allProductCategoryMap.get(productCategoryEntity.getProductCategoryId()) != null) {
                        productCategoryName = allProductCategoryMap.get(productCategoryEntity.getProductCategoryId()).getName();
                        if (productCategoryName.length() > 2) {
                            productCategoryName = StringUtils.left(productCategoryName, 2);
                        }
                    }
                }
                productCategoryEntity.setProductCategoryName(productCategoryName);
                productCategoryEntity.setWithoutServicePoint(servicePointStreetTotal - productCategoryEntity.getFrequentServicePoint() - productCategoryEntity.getTrialServicePoint());
            }
        }

        if (servicePointStreetTotal != 0) {
            frequentQtyRate = numberFormat.format((float) servicePointStreetQty.getFrequentServicePoint() / servicePointStreetTotal * 100);
            trialQtyRate = numberFormat.format((float) servicePointStreetQty.getTrialServicePoint() / servicePointStreetTotal * 100);
            withoutQtyRate = numberFormat.format((float) servicePointStreetQty.getWithoutServicePoint() / servicePointStreetTotal * 100);
            servicePointStreetQty.setFrequentServicePointRate(frequentQtyRate);
            servicePointStreetQty.setTrialServicePointRate(trialQtyRate);
            servicePointStreetQty.setWithoutServicePointRate(withoutQtyRate);

            autoPlanStreetRate = numberFormat.format((float) servicePointStreetAutoPlanQty.getAutoPlanStreet() / servicePointStreetTotal * 100);
            servicePointStreetAutoPlanQty.setAutoPlanStreetRate(autoPlanStreetRate);
        }
        map.put("servicePointStreetQty", servicePointStreetQty);
        map.put("servicePointStreetAutoPlanQty", servicePointStreetAutoPlanQty);
        map.put("servicePointProductCategoryQty", servicePointProductCategoryQty);

        return map;
    }

    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveServicePointStreetQtyToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteServicePointStreetQtyFromRptDB(beginDate);
                            saveServicePointStreetQtyToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointStreetQtyFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointStreetQtyService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    public void saveServicePointStreetQtyToRptDB(Date date) {
        List<RPTServicePointStreetQtyEntity> list = writeToIntermediateTable(date);

        for (RPTServicePointStreetQtyEntity entity : list) {
            servicePointStreetQtyMapper.insertServicePointStreetQtyData(entity);
        }
    }

    /**
     * 删除中间表中指定日期的数据
     */
    private void deleteServicePointStreetQtyFromRptDB(Date date) {
        if (date != null) {
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            servicePointStreetQtyMapper.deleteServicePointStreetQtyFromRptDB(systemId, startDate.getTime(), endDate.getTime());
        }
    }

    private List<RPTServicePointStreetQtyEntity> writeToIntermediateTable(Date date) {
        int systemId = RptCommonUtils.getSystemId();
        Date beginDate = DateUtils.getDateStart(date);

        List<NameValuePair<Integer, Long>> servicePointStreetList = msServicePointStreetService.findAllServicePointQtyForRPT();

        Map<Long, List<NameValuePair<Integer, Long>>> servicePointQtyCategory = msServicePointStreetService.findAllServicePointQtyCategoryForRPT();

        List<NameValuePair<Integer, Long>> servicePointAutoPlanList = msServicePointStreetService.findAllServicePointAutoPlanForRPT();



        RPTServicePointStreetQtyEntity entity = new RPTServicePointStreetQtyEntity();
        List<RPTServicePointStreetQtyEntity> list = Lists.newArrayList();
        entity.setSystemId(systemId);
        entity.setCreateDate(beginDate.getTime());
        entity.setProductCategoryId(0L);
        try {

            Map<Long,RPTArea> cacheTownAreaMap = areaCacheService.getAllTownMap();
            List<RPTArea> cacheTownAreaList = cacheTownAreaMap.values().stream().distinct().collect(Collectors.toList());

            Integer streetTotal = cacheTownAreaList.size();

            for (NameValuePair<Integer, Long> nameValuePair : servicePointStreetList) {
                if (nameValuePair.getName() == 10) {
                    entity.setTrialServicePoint(nameValuePair.getValue());
                } else if (nameValuePair.getName() == 20) {
                    entity.setFrequentServicePoint(nameValuePair.getValue());
                }
                entity.setServicePointStreet(streetTotal.longValue());
                entity.setWithoutServicePoint(streetTotal - entity.getTrialServicePoint() - entity.getFrequentServicePoint());
            }

            for (NameValuePair<Integer, Long> autoPlan : servicePointAutoPlanList) {
                if (autoPlan.getName() == 1) {
                    entity.setAutoPlanStreet(autoPlan.getValue());
                } else if (autoPlan.getName() == 2) {
                    entity.setAutoPlanWithoutFrequent(autoPlan.getValue());
                } else if (autoPlan.getName() == 3) {
                    entity.setFrequentWithoutAutoPlan(autoPlan.getValue());
                }
            }

            list.add(entity);


            for (Map.Entry<Long, List<NameValuePair<Integer, Long>>> entry : servicePointQtyCategory.entrySet()) {
                entity = new RPTServicePointStreetQtyEntity();

                entity.setProductCategoryId(entry.getKey());
                entity.setSystemId(systemId);
                entity.setCreateDate(beginDate.getTime());

                for (NameValuePair<Integer, Long> ignored : entry.getValue()) {
                    if (ignored.getName() == 10) {
                        entity.setTrialServicePoint(ignored.getValue());
                    } else if (ignored.getName() == 20) {
                        entity.setFrequentServicePoint(ignored.getValue());
                    }
                    entity.setWithoutServicePoint(streetTotal - entity.getTrialServicePoint() - entity.getFrequentServicePoint());
                }
                list.add(entity);
            }
        } catch (Exception e) {

            log.error("按品类网点街道数量写入失败ServicePointStreetQtyService.writeToIntermediateTable:{}", Exceptions.getStackTraceAsString(e));
        }

        return list;
    }
}
