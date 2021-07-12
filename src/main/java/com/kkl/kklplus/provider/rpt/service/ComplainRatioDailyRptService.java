package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTAreaCompletedDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTGradedOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.entity.rpt.mq.MQRPTUpdateOrderComplainMessage;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.*;
import com.kkl.kklplus.provider.rpt.mapper.ComplainRatioDailyRptMapper;
import com.kkl.kklplus.provider.rpt.mapper.GradedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.http.util.TextUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ComplainRatioDailyRptService  extends RptBaseService {

    @Resource
    private GradedOrderRptMapper gradedOrderRptMapper;

    @Autowired
    private ComplainRatioDailyRptMapper complainRatioDailyRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;


    /**
     * 获取省每日完工和投诉数据
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getProvinceCompletedOrderData(RPTGradedOrderSearch searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        searchCondition.setSystemId(systemId);
        List<RPTAreaCompletedDailyEntity> provinceGradedOrderData = gradedOrderRptMapper.getProvinceGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(),systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );
        List<RPTAreaCompletedDailyEntity> provinceComplainOrderData = complainRatioDailyRptMapper.getProvinceComplainOrderData(searchCondition);

        List<RPTAreaCompletedDailyEntity>  provinceServicePointBadList = complainRatioDailyRptMapper.getProvinceServicePointBadOrderData(searchCondition);

        Set<Long> salesIdSet = Sets.newHashSet();
        Map<Long, List<RPTAreaCompletedDailyEntity>> provinceComplainMap = provinceComplainOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getProvinceId));
        Iterator<Long> complainIter = provinceComplainMap.keySet().iterator();
        while (complainIter.hasNext()){
            salesIdSet.add(complainIter.next());
        }
        Map<Long, List<RPTAreaCompletedDailyEntity>> provinceCompletedMap = provinceGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getProvinceId));
        Iterator<Long> completedIter = provinceCompletedMap.keySet().iterator();
        while (completedIter.hasNext()){
            salesIdSet.add(completedIter.next());
        }
        Map<Long, List<RPTAreaCompletedDailyEntity>> provinceServicePointBadMap = provinceServicePointBadList.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getProvinceId));
        Iterator<Long> servicePointBadIter = provinceServicePointBadMap.keySet().iterator();
        while (servicePointBadIter.hasNext()){
            salesIdSet.add(servicePointBadIter.next());
        }
        Map<Long, RPTArea> areaMap = areaCacheService.getAllProvinceMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();
        RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity;
        for (Long provinceId :salesIdSet) {
            rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            List<RPTAreaCompletedDailyEntity>  provinceComplainList =  provinceComplainMap.get(provinceId);
            List<RPTAreaCompletedDailyEntity>  provinceCompletedList =  provinceCompletedMap.get(provinceId);
            List<RPTAreaCompletedDailyEntity>  servicePointBadList = provinceServicePointBadMap.get(provinceId);
            RPTArea rptArea = areaMap.get(provinceId);
            if (rptArea!=null){
                rptAreaCompletedDailyEntity.setProvinceId(rptArea.getId());
                rptAreaCompletedDailyEntity.setProvinceName(rptArea.getName());
            }
            Double completedTotal = 0.0;
            Double complainTotal = 0.0;
            Double servicePointTotal = 0.0;
            if(provinceCompletedList != null) {
                Class completedItemClass = rptAreaCompletedDailyEntity.getClass();
                for (RPTAreaCompletedDailyEntity entity : provinceCompletedList) {
                    completedTotal += Double.valueOf(entity.getCountSum());
                    String strSetDMethodName = "setD" + entity.getDayIndex();
                    try {
                        Method setDMethod = completedItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }

            if(provinceComplainList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : provinceComplainList){
                    complainTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setA" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }

            if(servicePointBadList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : servicePointBadList){
                    servicePointTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setC" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }
            rptAreaCompletedDailyEntity.setTotal(completedTotal);
            rptAreaCompletedDailyEntity.setTotalAmount(complainTotal);
            rptAreaCompletedDailyEntity.setEvaluateAmount(servicePointTotal);
            list.add(rptAreaCompletedDailyEntity);

        }
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setProvinceName("总计(单)");
        sumUp.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        sumUp.computeSumAndPerForAmount(list, 0, 0, sumUp, null);
        computeSumAndPerForCountOrAmount(list,sumUp);
        list.add(sumUp);
        return list;
    }

    /**
     * 获取市每日完工和投诉数据
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getCityCompletedOrderData(RPTGradedOrderSearch searchCondition) {
        int systemId = RptCommonUtils.getSystemId();
        searchCondition.setSystemId(systemId);
        List<RPTAreaCompletedDailyEntity> cityGradedOrderData = gradedOrderRptMapper.getCityGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(), systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );
        List<RPTAreaCompletedDailyEntity> cityComplainOrderData = complainRatioDailyRptMapper.getCityComplainOrderData(searchCondition);
        Map<Long, List<RPTAreaCompletedDailyEntity>> cityCompletedMap = cityGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCityId));

        List<RPTAreaCompletedDailyEntity>  cityServicePointBadList = complainRatioDailyRptMapper.getCityServicePointBadOrderData(searchCondition);

        Set<Long> salesIdSet = Sets.newHashSet();
        Iterator<Long> completedIter = cityCompletedMap.keySet().iterator();
        while (completedIter.hasNext()){
            salesIdSet.add(completedIter.next());
        }

        Map<Long, List<RPTAreaCompletedDailyEntity>> cityComplainMap = cityComplainOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCityId));
        Iterator<Long> complainIter = cityComplainMap.keySet().iterator();
        while (complainIter.hasNext()){
            salesIdSet.add(complainIter.next());
        }
        Map<Long, List<RPTAreaCompletedDailyEntity>> cityServicePointBadMap = cityServicePointBadList.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCityId));
        Iterator<Long> servicePointBadIter = cityServicePointBadMap.keySet().iterator();
        while (servicePointBadIter.hasNext()){
            salesIdSet.add(servicePointBadIter.next());
        }
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();
        RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity;
        for (Long cityId :salesIdSet) {
            rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            List<RPTAreaCompletedDailyEntity>  cityComplainList =  cityComplainMap.get(cityId);
            List<RPTAreaCompletedDailyEntity>  cityCompletedList =  cityCompletedMap.get(cityId);
            List<RPTAreaCompletedDailyEntity>  cityPointBadList = cityServicePointBadMap.get(cityId);
            RPTArea rptArea = areaMap.get(cityId);

            if (rptArea != null) {
                rptAreaCompletedDailyEntity.setCityId(rptArea.getId());
                rptAreaCompletedDailyEntity.setCityName(rptArea.getName());
                RPTArea parent = rptArea.getParent();
                    if (parent!=null){
                        RPTArea provinceArea = provinceMap.get(parent.getId());
                        if(provinceArea != null){
                            rptAreaCompletedDailyEntity.setProvinceId(parent.getId());
                            rptAreaCompletedDailyEntity.setProvinceName(provinceArea.getName());
                        }
                    }
            }
            Double completedTotal = 0.0;
            Double complainTotal = 0.0;
            Double servicePointTotal = 0.0;
            if(cityCompletedList != null) {
                Class completedItemClass = rptAreaCompletedDailyEntity.getClass();
                for (RPTAreaCompletedDailyEntity entity : cityCompletedList) {
                    completedTotal += Double.valueOf(entity.getCountSum());
                    String strSetDMethodName = "setD" + entity.getDayIndex();
                    try {
                        Method setDMethod = completedItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }

            if(cityComplainList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : cityComplainList){
                    complainTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setA" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }

            if(cityPointBadList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : cityPointBadList){
                    servicePointTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setC" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }
            rptAreaCompletedDailyEntity.setTotalAmount(complainTotal);
            rptAreaCompletedDailyEntity.setTotal(completedTotal);
            rptAreaCompletedDailyEntity.setEvaluateAmount(servicePointTotal);
            list.add(rptAreaCompletedDailyEntity);

        }
        list = list.stream().sorted(Comparator.comparing(RPTAreaCompletedDailyEntity::getProvinceId)).collect(Collectors.toList());
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setProvinceName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        RPTBaseDailyEntity.computeSumAndPerForAmount(list, 0, 0, sumUp, null);
        computeSumAndPerForCountOrAmount(list,sumUp);
        list.add(sumUp);
        return list;
    }

    /**
     * 获取区域每日完工和投诉数据
     *
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getAreaCompletedOrderData(RPTGradedOrderSearch searchCondition) {
        int systemId = RptCommonUtils.getSystemId();
        searchCondition.setSystemId(systemId);
        List<RPTAreaCompletedDailyEntity> areaGradedOrderData = gradedOrderRptMapper.getAreaGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(), systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );

        List<RPTAreaCompletedDailyEntity> areaComplainOrderData = complainRatioDailyRptMapper.getAreaComplainOrderData(searchCondition);
        Map<Long, List<RPTAreaCompletedDailyEntity>> areaCompletedMap = areaGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCountyId));

        List<RPTAreaCompletedDailyEntity>  areaServicePointBadList = complainRatioDailyRptMapper.getAreaServicePointBadOrderData(searchCondition);

        Set<Long> areaIdSet = Sets.newHashSet();
        Iterator<Long> completedIter = areaCompletedMap.keySet().iterator();
        while (completedIter.hasNext()){
            areaIdSet.add(completedIter.next());
        }

        Map<Long, List<RPTAreaCompletedDailyEntity>> areaComplainMap = areaComplainOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCountyId));

        Iterator<Long> complainIter = areaComplainMap.keySet().iterator();
        while (complainIter.hasNext()){
            areaIdSet.add(complainIter.next());
        }

        Map<Long, List<RPTAreaCompletedDailyEntity>> areaServicePointBadMap = areaServicePointBadList.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCountyId));
        Iterator<Long> servicePointBadIter = areaServicePointBadMap.keySet().iterator();
        while (servicePointBadIter.hasNext()){
            areaIdSet.add(servicePointBadIter.next());
        }
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();
        RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity;
        for (Long areaId : areaIdSet) {
            rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            List<RPTAreaCompletedDailyEntity> areaComplainList =  areaComplainMap.get(areaId);
            List<RPTAreaCompletedDailyEntity> areaCompletedList =  areaCompletedMap.get(areaId);
            List<RPTAreaCompletedDailyEntity> areaPointBadList = areaServicePointBadMap.get(areaId);
            RPTArea rptArea  = areaMap.get(areaId);
            if (rptArea != null) {
                rptAreaCompletedDailyEntity.setCountyId(rptArea.getId());
                rptAreaCompletedDailyEntity.setCountyName(rptArea.getName());
                RPTArea city = rptArea.getParent();
                if (city!=null){
                    RPTArea cityArea = cityMap.get(city.getId());
                    if (cityArea!=null){
                        RPTArea parent = cityArea.getParent();
                        rptAreaCompletedDailyEntity.setCityId(cityArea.getId());
                        rptAreaCompletedDailyEntity.setCityName(cityArea.getName());
                        if (parent!=null){
                            RPTArea provinceArea = provinceMap.get(parent.getId());
                            if(provinceArea != null){
                                rptAreaCompletedDailyEntity.setProvinceId(parent.getId());
                                rptAreaCompletedDailyEntity.setProvinceName(provinceArea.getName());
                            }
                        }
                    }
                }
            }

            Double completedTotal = 0.0;
            Double complainTotal = 0.0;
            Double servicePointTotal = 0.0;
            if(areaCompletedList != null) {
                Class completedItemClass = rptAreaCompletedDailyEntity.getClass();
                for (RPTAreaCompletedDailyEntity entity : areaCompletedList) {
                    completedTotal += Double.valueOf(entity.getCountSum());
                    String strSetDMethodName = "setD" + entity.getDayIndex();
                    try {
                        Method setDMethod = completedItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }

            if(areaComplainList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : areaComplainList){
                    complainTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setA" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }

            if(areaPointBadList != null){
                Class complainItemClass = rptAreaCompletedDailyEntity.getClass();
                for(RPTAreaCompletedDailyEntity item : areaPointBadList){
                    servicePointTotal += Double.valueOf(item.getCountSum());
                    String strSetDMethodName = "setC" + item.getDayIndex();
                    try {
                        Method setDMethod = complainItemClass.getMethod(strSetDMethodName, Double.class);
                        setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(item.getCountSum()));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                }
            }
            rptAreaCompletedDailyEntity.setTotal(completedTotal);
            rptAreaCompletedDailyEntity.setTotalAmount(complainTotal);
            rptAreaCompletedDailyEntity.setEvaluateAmount(servicePointTotal);
            list.add(rptAreaCompletedDailyEntity);

        }
        list = list.stream().sorted(Comparator.comparing(RPTAreaCompletedDailyEntity::getProvinceId)
                .thenComparing(RPTAreaCompletedDailyEntity::getCityId)).collect(Collectors.toList());
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setCountyName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        RPTBaseDailyEntity.computeSumAndPerForAmount(list, 0, 0, sumUp, null);
        computeSumAndPerForCountOrAmount(list,sumUp);
        list.add(sumUp);
        return list;
    }

    public static void computeSumAndPerForCountOrAmount(List baseDailyReports, RPTAreaCompletedDailyEntity sumDailyReport) {

        Class sumDailyReportClass;
        if (sumDailyReport != null) {
            sumDailyReportClass = sumDailyReport.getClass();
            Iterator var8 = baseDailyReports.iterator();

            while(var8.hasNext()) {
                Object object = var8.next();
                RPTAreaCompletedDailyEntity item = (RPTAreaCompletedDailyEntity)object;
                Class itemClass = item.getClass();

                for(int i = 1; i < 32; ++i) {
                    String strGetMethodName = "getC" + i;
                    String strSetMethodName = "setC" + i;

                    try {
                        Method itemGetMethod = itemClass.getMethod(strGetMethodName);
                        Object itemGetD = itemGetMethod.invoke(item);
                        Method sumDailyReportClassGetMethod = sumDailyReportClass.getMethod(strGetMethodName);
                        Object sumDailyReportClassGetD = sumDailyReportClassGetMethod.invoke(sumDailyReport);
                        double dSum = 0.0D;
                        if (sumDailyReportClassGetD != null && itemGetD != null) {
                            dSum = Double.valueOf(sumDailyReportClassGetD.toString().trim()) + Double.valueOf(itemGetD.toString().trim());
                        }

                        Method sumDailyReportSetMethod = sumDailyReportClass.getMethod(strSetMethodName, Double.class);
                        sumDailyReportSetMethod.invoke(sumDailyReport, dSum);
                    } catch (Exception var23) {
                        var23.printStackTrace();
                    }
                }

                double dSumTotal = sumDailyReport.getEvaluateAmount() + item.getEvaluateAmount();
                sumDailyReport.setEvaluateAmount(dSumTotal);
            }
        }
    }



    /**
     *保存消息队列的工单到中间表
     */
    public void saveOrderComplainMQ(MQRPTUpdateOrderComplainMessage.MQOrderComplainMessage msg){
        if(msg.getId()<=0){
            throw new RuntimeException("投诉单id不能为空");
        }
        if(msg.getStatus() < 0){
            throw new RuntimeException("投诉单状态不能为空");
        }

        ComplainDailyEntity item = new ComplainDailyEntity();
        List<Integer> ids;
        List<Integer> itemIds;
        if(msg.getComplainDt() > 0){
            String quarter = QuarterUtils.getSeasonQuarter(new Date(msg.getComplainDt()));
            item.setQuarter(quarter);
        }
        item.setComplainId(msg.getId());

        if (msg.getJudgeObject() > 0) {
            ids = BytesUtils.intToIntegerList(msg.getJudgeObject());
            if (ids != null && ids.size() > 0) {
                for (Integer id : ids) {
                    if (id == 4 || id == 5) {
                        item.setComplainStatus(1);
                    }
                    if (id == 1) {
                        if (msg.getJudgeItem() > 0) {
                            itemIds = BytesUtils.intToIntegerList(msg.getJudgeItem());
                            if (itemIds != null && itemIds.size() > 0) {
                                for (Integer itemId : itemIds) {
                                    if (itemId == 1) {
                                        item.setEvaluateStatus(1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if (msg.getStatus() == 4) {
            item.setComplainStatus(1);
        }
        int systemId = RptCommonUtils.getSystemId();
        updateComplainOrderToRptMQ(item,systemId);

    }

    public List<ComplainDailyEntity>  getComplainOrderList(Date startDate,  Date endDate){
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        int day = Integer.valueOf(DateUtils.getDay(startDate));
        List<ComplainDailyEntity> complainOrderList = complainRatioDailyRptMapper.getComplainDailyList(startDate, endDate);

        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        List<Integer> ids;
        List<Integer> itemIds;
        for(ComplainDailyEntity item :complainOrderList) {
            if (item.getComplainDate() != null) {
                item.setComplainDt(item.getComplainDate().getTime());
            }
            item.setDayIndex(day);
            item.setSystemId(systemId);
            item.setQuarter(quarter);

            RPTArea rptArea = areaMap.get(item.getCountyId());
            if (rptArea != null) {
                RPTArea city = rptArea.getParent();
                if (city != null) {
                    item.setCityId(city.getId());
                    RPTArea cityArea = cityMap.get(city.getId());
                    if (cityArea != null) {
                        RPTArea parent = cityArea.getParent();
                        if (parent != null) {
                            item.setProvinceId(parent.getId());
                        }
                    }
                }
            }

            if (item.getJudgeObject() > 0) {
                ids = BytesUtils.intToIntegerList(item.getJudgeObject());
                if (ids != null && ids.size() > 0) {
                    for (Integer id : ids) {
                        if (id == 4 || id == 5) {
                            item.setComplainStatus(1);
                        }
                        if (id == 1) {
                            if (item.getJudgeItem() > 0) {
                                itemIds = BytesUtils.intToIntegerList(item.getJudgeItem());
                                if (itemIds != null && itemIds.size() > 0) {
                                    for (Integer itemId : itemIds) {
                                        if (itemId == 1) {
                                            item.setEvaluateStatus(1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (item.getStatus() == 4) {
                item.setComplainStatus(1);
            }
        }
           return complainOrderList;
    }


    private Map<Long, Long> getComplainIdMap(int systemId, Date startDate, Date endDate,String quarter) {
        List<LongTwoTuple> tuples = complainRatioDailyRptMapper.getComplainOrderIds(systemId,startDate.getTime(),endDate.getTime(),quarter);
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

    public void updateComplainOrderToRptDB(Date date) {
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<ComplainDailyEntity> list = getComplainOrderList(startDate, endDate);
        if (!list.isEmpty()) {
            int systemId = RptCommonUtils.getSystemId();
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Map<Long, Long> idMap = getComplainIdMap(systemId, startDate, endDate,quarter);
            Long primaryKeyId;
            for (ComplainDailyEntity item : list) {
                primaryKeyId = idMap.get(item.getComplainId());
                if (primaryKeyId == null || primaryKeyId == 0) {
                    complainRatioDailyRptMapper.insertComplainDaily(item);
                }
            }
        }
    }


    public void updateComplainOrderToRptMQ(ComplainDailyEntity complainOrderOfMQ , Integer systemId) {
            complainOrderOfMQ.setSystemId(systemId);
            complainRatioDailyRptMapper.updateComplainDaily(complainOrderOfMQ);
    }




    public void saveComplainOrderRptDB(Date date){
        if(date != null){
            Date startDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            List <ComplainDailyEntity> list = getComplainOrderList(startDate,endDate);
            for(ComplainDailyEntity item : list){
                complainRatioDailyRptMapper.insertComplainDaily(item);
            }
        }
    }



    private void deleteComplainOrderRptDB(Date date){
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            String quarter = QuarterUtils.getSeasonQuarter(date);
            complainRatioDailyRptMapper.deleteComplainData(beginDate.getTime(), endDate.getTime(), systemId,quarter);
        }
    }




    public boolean rebuildMiddleTableComplainData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveComplainOrderRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            updateComplainOrderToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteComplainOrderRptDB(beginDate);
                            saveComplainOrderRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteComplainOrderRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ComplainRatioDailyRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;

    }

    /**
     * 检查区域每日完工报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition != null) {
            if (new Date().getTime() < searchCondition.getBeginDate()) {
                return false;
            }
            searchCondition.setSystemId(systemId);
            Integer rowCount = gradedOrderRptMapper.hasAreaCompletedOrderReportData(searchCondition.getBeginDate(),
                    searchCondition.getEndDate(),searchCondition.getAreaType(),searchCondition.getAreaId(),systemId,searchCondition.getCustomerId(),
                    searchCondition.getQuarter(),searchCondition.getProductCategoryIds(),searchCondition.getQuarters());
            Integer count = complainRatioDailyRptMapper.hasAreaComplainCompletedReportData(searchCondition);
            result = rowCount+count > 0;
        }
        return result;
    }

    /**
     * 创建省每日完工报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportProvinceComplainCompletedRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            List<RPTAreaCompletedDailyEntity> list = getProvinceCompletedOrderData(searchCondition);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getBeginDate()));
            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days * 5 + 5));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1,  days*5));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工和投诉(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1+days*5, 5+days * 5));
            ExportExcel.createCell(headFirstRow, days*5 + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                xSheet.addMergedRegion(new CellRangeAddress(2, 2, dayIndex * 5 - 4,  dayIndex * 5));
                ExportExcel.createCell(headSecondRow, dayIndex * 5- 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex);
            }

            Row headThirdRow = xSheet.createRow(rowIndex++);
            headThirdRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days+1; dayIndex++) {
                ExportExcel.createCell(headThirdRow, dayIndex * 5-4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉比率");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评");
                ExportExcel.createCell(headThirdRow, dayIndex * 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评比率");

            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
            NumberFormat numberFormat = NumberFormat.getInstance();
            numberFormat.setMaximumFractionDigits(2);
            if (list != null && list.size() > 0) {
                for (RPTAreaCompletedDailyEntity entity : list) {
                    index ++;
                    if (list.size()>index){
                        dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int pColumnIndex = 0;
                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                        Class provinceItemClass = entity.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getD" + dayIndex;
                            String strGetAMethodName = "getA" + dayIndex;
                            String strGetCMethodName = "getC" + dayIndex;

                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Method getAMethod = provinceItemClass.getMethod(strGetAMethodName);
                            Method getCMethod = provinceItemClass.getMethod(strGetCMethodName);

                            Object objGetD = getDMethod.invoke(entity);
                            Object objGetA = getAMethod.invoke(entity);
                            Object objGetC = getCMethod.invoke(entity);

                            Double dSum = StringUtils.toDouble(objGetD);
                            Double aSum = StringUtils.toDouble(objGetA);
                            Double cSum = StringUtils.toDouble(objGetC);

                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dSum);
                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aSum);
                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(aSum *100 ): numberFormat.format(aSum/dSum *100))+"%");
                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cSum);
                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(cSum *100) : numberFormat.format(cSum/dSum *100))+"%");
                        }
                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotal());
                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotalAmount());

                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (entity.getTotal() == 0 ?
                                numberFormat.format(entity.getTotalAmount()*100) : numberFormat.format(entity.getTotalAmount()/entity.getTotal()*100)) + "%");

                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getEvaluateAmount());

                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (entity.getTotal() == 0 ?
                                numberFormat.format(entity.getEvaluateAmount()*100) : numberFormat.format(entity.getEvaluateAmount()/entity.getTotal()*100)) + "%");

                    }else {
                        //读取总计
                        dataRow = xSheet.createRow(rowIndex);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 1;

                        dataCell = ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                        dataCell.setCellValue(null == entity.getProvinceName() ? "" : entity.getProvinceName());

                        Class totalClass = entity.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getD" + dayIndex;
                            String strGetAMethodName = "getA" + dayIndex;
                            String strGetCMethodName = "getC" + dayIndex;

                            Method getDMethod = totalClass.getMethod(strGetDMethodName);
                            Method getAMethod = totalClass.getMethod(strGetAMethodName);
                            Method getCMethod = totalClass.getMethod(strGetCMethodName);

                            Object objGetC = getCMethod.invoke(entity);
                            Object objGetD = getDMethod.invoke(entity);
                            Object objGetA = getAMethod.invoke(entity);

                            Double dSum = StringUtils.toDouble(objGetD);
                            Double aSum = StringUtils.toDouble(objGetA);
                            Double cSum = StringUtils.toDouble(objGetC);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dSum);
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,aSum);
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(aSum *100 ): numberFormat.format(aSum/dSum *100))+"%");
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cSum);
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(cSum *100) : numberFormat.format(cSum/dSum *100))+"%");
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotal());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotalAmount());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (entity.getTotal() == 0 ?
                                numberFormat.format(entity.getTotalAmount()*100) : numberFormat.format(entity.getTotalAmount()/entity.getTotal()*100)) + "%");

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getEvaluateAmount());

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (entity.getTotal() == 0 ?
                                numberFormat.format(entity.getEvaluateAmount()*100) : numberFormat.format(entity.getEvaluateAmount()/entity.getTotal()*100)) + "%");
                    }

                }
            }
        }
        catch (Exception e) {
            log.error("省每日完工和投诉写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


    /**
     * 创建市每日完工报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportCityComplainCompletedRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            List<RPTAreaCompletedDailyEntity> cityList = getCityCompletedOrderData(searchCondition);
            List<RPTAreaCompletedDailyEntity> provinceList = getProvinceCompletedOrderData(searchCondition);
            RPTAreaCompletedDailyEntity sumAOPD = provinceList.get(provinceList.size()-1);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getBeginDate()));
            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days*5 + 6));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, days*5+1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工和投诉(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days*5 + 2, days*5 + 6));
            ExportExcel.createCell(headFirstRow, days*5 + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                xSheet.addMergedRegion(new CellRangeAddress(2, 2, dayIndex * 5-3,  dayIndex * 5+1));
                ExportExcel.createCell(headSecondRow, dayIndex * 5-3 , xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex);
            }

            Row headThirdRow = xSheet.createRow(rowIndex++);
            headThirdRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days+1; dayIndex++) {
                ExportExcel.createCell(headThirdRow, dayIndex * 5-3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉比率");
                ExportExcel.createCell(headThirdRow, dayIndex * 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评");
                ExportExcel.createCell(headThirdRow, dayIndex * 5+1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评比率");
            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
            NumberFormat numberFormat = NumberFormat.getInstance();
            numberFormat.setMaximumFractionDigits(2);
            if (provinceList.size()>0) {
                for (int i = 0; i < provinceList.size()-1; i++) {
                    RPTAreaCompletedDailyEntity province = provinceList.get(i);
                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int pColumnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == province.getProvinceName() ? "" : province.getProvinceName());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    Class provinceClass = province.getClass();
                    pColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, province, pColumnIndex, provinceClass,numberFormat);

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotal());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotalAmount());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (province.getTotal() == 0 ?
                            numberFormat.format(province.getTotalAmount()*100) : numberFormat.format(province.getTotalAmount()/province.getTotal()*100)) + "%");

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getEvaluateAmount());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (province.getTotal() == 0 ?
                            numberFormat.format(province.getEvaluateAmount()*100) : numberFormat.format(province.getEvaluateAmount()/province.getTotal()*100)) + "%");

                    if (cityList != null && cityList.size() > 0) {
                        for (int j = 0; j < cityList.size() - 1; j++) {
                            RPTAreaCompletedDailyEntity city = cityList.get(j);
                            if (city.getProvinceId().intValue() == province.getProvinceId().intValue()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                int cColumnIndex = 0;
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());
                                Class cityClass = city.getClass();
                                cColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, city, cColumnIndex, cityClass,numberFormat);
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotal());
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotalAmount());

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (city.getTotal() == 0 ?
                                        numberFormat.format(province.getTotalAmount()*100) : numberFormat.format(city.getTotalAmount()/city.getTotal()*100)) + "%");

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getEvaluateAmount());

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (city.getTotal() == 0 ?
                                        numberFormat.format(city.getEvaluateAmount()*100) : numberFormat.format(city.getEvaluateAmount()/city.getTotal()*100)) + "%");
                            }
                        }

                    }
                }
                //读取总计
                dataRow = xSheet.createRow(rowIndex);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                int columnIndex = 2;

                xSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 1));
                dataCell = ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                dataCell.setCellValue(null == sumAOPD.getProvinceName() ? "" : sumAOPD.getProvinceName());

                Class totalClass = sumAOPD.getClass();
                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                    String strGetDMethodName = "getD" + dayIndex;
                    String strGetAMethodName = "getA" + dayIndex;
                    String strGetCMethodName = "getC" + dayIndex;

                    Method getDMethod = totalClass.getMethod(strGetDMethodName);
                    Method getAMethod = totalClass.getMethod(strGetAMethodName);
                    Method getCMethod = totalClass.getMethod(strGetCMethodName);

                    Object objGetC = getCMethod.invoke(sumAOPD);
                    Object objGetD = getDMethod.invoke(sumAOPD);
                    Object objGetA = getAMethod.invoke(sumAOPD);

                    Double dSum = StringUtils.toDouble(objGetD);
                    Double aSum = StringUtils.toDouble(objGetA);
                    Double cSum = StringUtils.toDouble(objGetC);

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,aSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(aSum *100 ): numberFormat.format(aSum/dSum *100))+"%");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(cSum *100) : numberFormat.format(cSum/dSum *100))+"%");
                }

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getTotal());
                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getTotalAmount());

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (sumAOPD.getTotal() == 0 ?
                        numberFormat.format(sumAOPD.getTotalAmount()*100) : numberFormat.format(sumAOPD.getTotalAmount()/sumAOPD.getTotal()*100)) + "%");

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getEvaluateAmount());

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (sumAOPD.getTotal() == 0 ?
                        numberFormat.format(sumAOPD.getEvaluateAmount()*100) : numberFormat.format(sumAOPD.getEvaluateAmount()/sumAOPD.getTotal()*100)) + "%");
            }
        }
        catch (Exception e) {
            log.error("市每日完工和投诉写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


    /**
     *创建区县每日完工和投诉报表
     */
    public SXSSFWorkbook exportCountyComplainCompletedRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            List<RPTAreaCompletedDailyEntity> areaList = getAreaCompletedOrderData(searchCondition);
            List<RPTAreaCompletedDailyEntity> cityList = getCityCompletedOrderData(searchCondition);
            List<RPTAreaCompletedDailyEntity> provinceList = getProvinceCompletedOrderData(searchCondition);

            RPTAreaCompletedDailyEntity sumAOPD = provinceList.get(provinceList.size()-1);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getBeginDate()));
            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days*5 + 7));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, days*5+2 ));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工和投诉(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days*5 + 3, days*5 + 7));
            ExportExcel.createCell(headFirstRow, days*5 + 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");
            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                xSheet.addMergedRegion(new CellRangeAddress(2, 2, dayIndex * 5-2,  dayIndex * 5+2));
                ExportExcel.createCell(headSecondRow, dayIndex * 5-2 , xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex);
            }

            Row headThirdRow = xSheet.createRow(rowIndex++);
            headThirdRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days+1; dayIndex++) {
                ExportExcel.createCell(headThirdRow, dayIndex * 5-2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成");
                ExportExcel.createCell(headThirdRow, dayIndex * 5-1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉");
                ExportExcel.createCell(headThirdRow, dayIndex * 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "有效投诉比率");
                ExportExcel.createCell(headThirdRow, dayIndex * 5+1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评");
                ExportExcel.createCell(headThirdRow, dayIndex * 5+2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点中差评比率");
            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
            NumberFormat numberFormat = NumberFormat.getInstance();
            numberFormat.setMaximumFractionDigits(2);
            if (provinceList.size()>0) {
                for (int i = 0; i < provinceList.size()-1; i++) {
                    RPTAreaCompletedDailyEntity province = provinceList.get(i);
                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int pColumnIndex = 0;
                    dataCell = ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == province.getProvinceName() ? "" : province.getProvinceName());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                    Class provinceClass = province.getClass();
                    pColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, province, pColumnIndex, provinceClass,numberFormat);
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotal());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotalAmount());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (province.getTotal() == 0 ?
                            numberFormat.format(province.getTotalAmount()*100) : numberFormat.format(province.getTotalAmount()/province.getTotal()*100)) + "%");

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getEvaluateAmount());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (province.getTotal() == 0 ?
                            numberFormat.format(province.getEvaluateAmount()*100) : numberFormat.format(province.getEvaluateAmount()/province.getTotal()*100)) + "%");


                    if (cityList != null && cityList.size() > 0) {
                        for (int j = 0; j < cityList.size() - 1; j++) {
                            RPTAreaCompletedDailyEntity city = cityList.get(j);
                            if (city.getProvinceId().intValue() == province.getProvinceId().intValue()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                int cColumnIndex = 0;
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                                Class cityClass = city.getClass();
                                cColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, city, cColumnIndex, cityClass,numberFormat);
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotal());
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotalAmount());

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (city.getTotal() == 0 ?
                                        numberFormat.format(province.getTotalAmount()*100) : numberFormat.format(city.getTotalAmount()/city.getTotal()*100)) + "%");

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getEvaluateAmount());

                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (city.getTotal() == 0 ?
                                        numberFormat.format(city.getEvaluateAmount()*100) : numberFormat.format(city.getEvaluateAmount()/city.getTotal()*100)) + "%");
                                //循环读取市下面的区
                                if (areaList != null && areaList.size() > 0) {
                                    for (int k = 0; k < areaList.size() - 1; k++) {
                                        RPTAreaCompletedDailyEntity area = areaList.get(k);
                                        if (area.getCityId().intValue() == city.getCityId().intValue()) {
                                            dataRow = xSheet.createRow(rowIndex++);
                                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                            int aColumnIndex = 0;

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                            dataCell = ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                            dataCell.setCellValue(null == area.getCountyName() ? "" : area.getCountyName());
                                            Class areaClass = area.getClass();

                                            aColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, area, aColumnIndex, areaClass,numberFormat);
                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, area.getTotal());
                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, area.getTotalAmount());

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (area.getTotal() == 0 ?
                                                    numberFormat.format(province.getTotalAmount()*100) : numberFormat.format(area.getTotalAmount()/area.getTotal()*100)) + "%");

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, area.getEvaluateAmount());

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (area.getTotal() == 0 ?
                                                    numberFormat.format(area.getEvaluateAmount()*100) : numberFormat.format(area.getEvaluateAmount()/area.getTotal()*100)) + "%");
                                        }
                                    }
                                }
                            }
                        }

                    }


                }
                //读取总计
                dataRow = xSheet.createRow(rowIndex);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                int columnIndex = 3;

                xSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
                dataCell = ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                dataCell.setCellValue(null == sumAOPD.getProvinceName() ? "" : sumAOPD.getProvinceName());


                Class totalClass = sumAOPD.getClass();

                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                    String strGetDMethodName = "getD" + dayIndex;
                    String strGetAMethodName = "getA" + dayIndex;
                    String strGetCMethodName = "getC" + dayIndex;

                    Method getDMethod = totalClass.getMethod(strGetDMethodName);
                    Method getAMethod = totalClass.getMethod(strGetAMethodName);
                    Method getCMethod = totalClass.getMethod(strGetCMethodName);

                    Object objGetC = getCMethod.invoke(sumAOPD);
                    Object objGetD = getDMethod.invoke(sumAOPD);
                    Object objGetA = getAMethod.invoke(sumAOPD);

                    Double dSum = StringUtils.toDouble(objGetD);
                    Double aSum = StringUtils.toDouble(objGetA);
                    Double cSum = StringUtils.toDouble(objGetC);

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,aSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(aSum *100 ): numberFormat.format(aSum/dSum *100))+"%");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cSum);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(cSum *100) : numberFormat.format(cSum/dSum *100))+"%");
                }
                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getTotal());
                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getTotalAmount());

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (sumAOPD.getTotal() == 0 ?
                        numberFormat.format(sumAOPD.getTotalAmount()*100) : numberFormat.format(sumAOPD.getTotalAmount()/sumAOPD.getTotal()*100)) + "%");

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, sumAOPD.getEvaluateAmount());

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (sumAOPD.getTotal() == 0 ?
                        numberFormat.format(sumAOPD.getEvaluateAmount()*100) : numberFormat.format(sumAOPD.getEvaluateAmount()/sumAOPD.getTotal()*100)) + "%");


            }

        }
        catch (Exception e) {
            log.error("区每日完工和投诉表写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

    private int writeDailyPlanOrders(int days, Map<String, CellStyle> xStyle, Row dataRow, RPTAreaCompletedDailyEntity entity, int pColumnIndex, Class provinceClass, NumberFormat numberFormat) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException
    {
        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
            String strGetDMethodName = "getD" + dayIndex;
            String strGetAMethodName = "getA" + dayIndex;
            String strGetCMethodName = "getC" + dayIndex;

            Method getDMethod = provinceClass.getMethod(strGetDMethodName);
            Method getAMethod = provinceClass.getMethod(strGetAMethodName);
            Method getCMethod = provinceClass.getMethod(strGetCMethodName);

            Object objGetD = getDMethod.invoke(entity);
            Object objGetA = getAMethod.invoke(entity);
            Object objGetC = getCMethod.invoke(entity);

            Double dSum = StringUtils.toDouble(objGetD);
            Double aSum = StringUtils.toDouble(objGetA);
            Double cSum = StringUtils.toDouble(objGetC);

            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dSum);
            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aSum);
            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(aSum *100 ): numberFormat.format(aSum/dSum *100))+"%");
            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cSum);
            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (dSum == 0? numberFormat.format(cSum *100) : numberFormat.format(cSum/dSum *100))+"%");

        }
        return pColumnIndex;
    }


}
