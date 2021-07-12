package com.kkl.kklplus.provider.rpt.chart.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.provider.rpt.chart.entity.RPTServicePointQtyEntity;
import com.kkl.kklplus.provider.rpt.chart.mapper.ServicePointQtyStatisticsMapper;
import com.kkl.kklplus.provider.rpt.chart.ms.md.service.MSServicePointQtyService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
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
public class ServicePointQtyStatisticsService {

    @Resource
    private ServicePointQtyStatisticsMapper servicePointQtyStatisticsMapper;

    @Autowired
    private MSServicePointQtyService servicePointQtyService;


    public Map<String, Object> getServicePointQtyData(RPTDataDrawingListSearch search) {
        Map<String, Object> map = new HashMap<>();
        int systemId = RptCommonUtils.getSystemId();
        Long endDate = DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime();
        Long startDate = DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime();
        Long productCategoryId = 0L;
        RPTServicePointQtyEntity entity = servicePointQtyStatisticsMapper.getServicePointQty(systemId, startDate, endDate, productCategoryId);
        if (entity == null) {
            return map;
        }
        Map<Long, RPTProductCategory> allProductCategoryMap = MDUtils.getAllProductCategoryMap();
        List<Long> productCategoryIds = Lists.newArrayList(allProductCategoryMap.keySet());
        List<RPTServicePointQtyEntity> servicePointProductCategoryQty = servicePointQtyStatisticsMapper.getServicePointProductCategoryQty(systemId, startDate, endDate, productCategoryIds);
        List<RPTServicePointQtyEntity> servicePointAutoPlanQty = servicePointQtyStatisticsMapper.getServicePointAutoPlanQty(systemId, startDate, endDate, productCategoryIds);

        long frequentQty = entity.getFrequentQty();
        long trialQty = entity.getTrialQty();

        long servicePointTotalQty = frequentQty + trialQty;


        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        String productCategoryName = "";
        long servicePointProductCategoryTotalQty = 0;
        String totalRate;
        for (RPTServicePointQtyEntity servicePointProductCategory : servicePointProductCategoryQty) {
            if (servicePointProductCategory.getProductCategoryId() != 0) {//服务品类可能为0  获取不到服务品类名称
                if(!allProductCategoryMap.isEmpty()){
                    if(allProductCategoryMap.get(servicePointProductCategory.getProductCategoryId()) != null){
                        productCategoryName = allProductCategoryMap.get(servicePointProductCategory.getProductCategoryId()).getName();
                        if (productCategoryName.length() > 2) {
                            productCategoryName = StringUtils.left(productCategoryName, 2);
                        }
                    }

                }

                servicePointProductCategory.setProductCategoryName(productCategoryName);
                servicePointProductCategoryTotalQty = servicePointProductCategory.getTotal();
                if (servicePointTotalQty != 0) {
                    totalRate = numberFormat.format((float) servicePointProductCategoryTotalQty / servicePointTotalQty * 100);
                    servicePointProductCategory.setTotalRate(totalRate);
                }
            }
        }
        productCategoryName = "";
        for (RPTServicePointQtyEntity servicePointAutoPlanEntity : servicePointAutoPlanQty) {
            if (servicePointAutoPlanEntity.getProductCategoryId() != 0) {   //服务品类可能为0  获取不到服务品类名称
                if(!allProductCategoryMap.isEmpty()){
                    if(allProductCategoryMap.get(servicePointAutoPlanEntity.getProductCategoryId()) != null) {
                        productCategoryName = allProductCategoryMap.get(servicePointAutoPlanEntity.getProductCategoryId()).getName();
                        if (productCategoryName.length() > 2) {
                            productCategoryName = StringUtils.left(productCategoryName, 2);
                        }
                    }
                }
                servicePointAutoPlanEntity.setProductCategoryName(productCategoryName);
                if (servicePointProductCategoryTotalQty != 0) {
                    totalRate = numberFormat.format((float) servicePointAutoPlanEntity.getAutoPlanQty() / servicePointProductCategoryTotalQty * 100);
                    servicePointAutoPlanEntity.setTotalRate(totalRate);
                }

            }
        }

        String frequentQtyRate = "";
        String trialQtyRate = "";


        if (servicePointTotalQty != 0) {
            frequentQtyRate = numberFormat.format((float) frequentQty / servicePointTotalQty * 100);
            trialQtyRate = numberFormat.format((float) trialQty / servicePointTotalQty * 100);
        }


        map.put("frequentQtyRate", frequentQtyRate);
        map.put("trialQtyRate", trialQtyRate);
        map.put("frequentQty", frequentQty);
        map.put("trialQty", trialQty);
        map.put("servicePointProductCategoryQty", servicePointProductCategoryQty);
        map.put("servicePointAutoPlanQty", servicePointAutoPlanQty);

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
                            saveServicePointQtyToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteServicePointQtyFromRptDB(beginDate);
                            saveServicePointQtyToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointQtyFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointQtyStatisticsService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    public void saveServicePointQtyToRptDB(Date date) {
        List<RPTServicePointQtyEntity> list = writeToIntermediateTable(date);

        for (RPTServicePointQtyEntity entity : list) {
            servicePointQtyStatisticsMapper.insertServicePointQtyData(entity);
        }

    }
    /**
     * 删除中间表中指定日期的数据
     */
    private void deleteServicePointQtyFromRptDB(Date date) {
        if (date != null) {
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            servicePointQtyStatisticsMapper.deleteServicePointQtyFromRptDB(systemId, startDate.getTime(), endDate.getTime());
        }
    }

    private List<RPTServicePointQtyEntity> writeToIntermediateTable(Date date) {
        int systemId = RptCommonUtils.getSystemId();
        Date beginDate = DateUtils.getDateStart(date);
        List<NameValuePair<Integer, Long>> servicePointQtyList = servicePointQtyService.findAllServicePointQtyForRPT();

        Map<Long, List<NameValuePair<Integer, Integer>>> qtyListMap = servicePointQtyService.findAllServicePointQtyCategoryForRPT();

        List<NameValuePair<Long, Long>> servicePointAutoPlanList = servicePointQtyService.findAllServicePointAutoPlanForRPT();
        Map<Long, Long> autoPlanMap = new HashMap<>();
        if(!servicePointAutoPlanList.isEmpty()){
            autoPlanMap = servicePointAutoPlanList.stream().collect(Collectors.toMap(NameValuePair::getName, NameValuePair::getValue));
        }
        RPTServicePointQtyEntity entity = new RPTServicePointQtyEntity();
        List<RPTServicePointQtyEntity> list = Lists.newArrayList();
        entity.setSystemId(systemId);
        entity.setCreateDate(beginDate.getTime());
        entity.setProductCategoryId(0L);
        for (NameValuePair<Integer, Long> nameValuePair : servicePointQtyList) {
            if (nameValuePair.getName() == 10) {
                entity.setTrialQty(nameValuePair.getValue());
            } else if (nameValuePair.getName() == 20) {
                entity.setFrequentQty(nameValuePair.getValue());
            }
        }
        list.add(entity);

        long autoPlanQty;
        for (Map.Entry<Long, List<NameValuePair<Integer, Integer>>> entry : qtyListMap.entrySet()) {
            entity = new RPTServicePointQtyEntity();
            autoPlanQty = 0;
            entity.setProductCategoryId(entry.getKey());
            entity.setSystemId(systemId);
            entity.setCreateDate(beginDate.getTime());
            if (!autoPlanMap.isEmpty()) {
                if(autoPlanMap.get(entry.getKey()) != null){
                    autoPlanQty = autoPlanMap.get(entry.getKey());
                }

            }
            entity.setAutoPlanQty(autoPlanQty);

            for (NameValuePair<Integer, Integer> ignored : entry.getValue()) {
                if (ignored.getName() == 10) {
                    entity.setTrialQty(ignored.getValue().longValue());
                } else if (ignored.getName() == 20) {

                    entity.setFrequentQty(ignored.getValue().longValue());
                }
            }
            list.add(entity);
        }

        return list;
    }
}
