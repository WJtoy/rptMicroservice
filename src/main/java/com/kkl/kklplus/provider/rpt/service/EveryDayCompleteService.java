package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;


import com.kkl.kklplus.entity.rpt.RPTAreaOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteEntity;
import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteSearch;
import com.kkl.kklplus.entity.rpt.search.RPTAreaOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;


import com.kkl.kklplus.provider.rpt.mapper.CustomerEveryDayCompleteMapper;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
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

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Auther wj
 * @Date 2021/5/21 17:19
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class EveryDayCompleteService  extends RptBaseService{

    @Autowired
    private CustomerEveryDayCompleteMapper customerEveryDayCompleteMapper;

    @Autowired
    private AreaCacheService areaCacheService;


    public Map<String, List<RPTEveryDayCompleteEntity>> getEveryDayComplete(RPTEveryDayCompleteSearch search){
        Date endDate = DateUtils.getEndOfDay(new Date(search.getEndDate()));
        Date startDate  = DateUtils.getStartOfDay(endDate);

        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        List<String> quarters = getQuarters(endDate);
        search.setQuarter(quarter);
        search.setStartDate(startDate.getTime());
        search.setEndDate(endDate.getTime());
        Map<String, List<RPTEveryDayCompleteEntity>> map = Maps.newHashMap();
        List<RPTEveryDayCompleteEntity> proList = customerEveryDayCompleteMapper.hasProvinceCompleteOrderData(search);//省下单
        List<RPTEveryDayCompleteEntity> rptList = customerEveryDayCompleteMapper.hasCompleteOrderData(search);  //市下单
        List<RPTEveryDayCompleteEntity> rptList1 = customerEveryDayCompleteMapper.hasCompleteRateData(search);  //完工单
        //未完工单
        List<RPTEveryDayCompleteEntity> unCompleteOrderData = customerEveryDayCompleteMapper.unCompleteOrderData(endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarters);

        startDate = DateUtils.addDays(startDate, -2);
        search.setStartDate(startDate.getTime());
        Date endDateNew = DateUtils.addDays(endDate , -2);
        search.setEndDate(endDateNew.getTime());
        String endQuarter = QuarterUtils.getSeasonQuarter(endDateNew);
        search.setQuarter(endQuarter);
        List<String> quarterList = QuarterUtils.getQuarters(startDate, endDate);
        List<RPTEveryDayCompleteEntity> planList48 = customerEveryDayCompleteMapper.hasCompleteOrderData(search);// 48小时前下的单
        List<RPTEveryDayCompleteEntity> rptList3 = customerEveryDayCompleteMapper.hasArrivalCompleteRateData(endDateNew,startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);


        startDate = DateUtils.addDays(startDate, -1);
        search.setStartDate(startDate.getTime());
        endDateNew = DateUtils.addDays(endDate , -3);
        search.setEndDate(endDateNew.getTime());
        endQuarter = QuarterUtils.getSeasonQuarter(endDateNew);
        search.setQuarter(endQuarter);
        quarterList = QuarterUtils.getQuarters(startDate, endDate);
        List<RPTEveryDayCompleteEntity> planList72 = customerEveryDayCompleteMapper.hasCompleteOrderData(search);// 72小时前下的单
        List<RPTEveryDayCompleteEntity> rptList2 = customerEveryDayCompleteMapper.hasCompleteRate72Data(endDateNew, startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);
        List<RPTEveryDayCompleteEntity> arrivalOrderData72 = customerEveryDayCompleteMapper.arrivalOrderData72(startDate,endDateNew,search.getAreaType(),search.getAreaId(),search.getCustomerId(),search.getQuarter());
        List<RPTEveryDayCompleteEntity> unArrivalOrderData72 = customerEveryDayCompleteMapper.unArrivalOrderData72(startDate,endDateNew,search.getAreaType(),search.getAreaId(),search.getCustomerId(),search.getQuarter());
        List<RPTEveryDayCompleteEntity> hasArrivalCompleteRateData72 =customerEveryDayCompleteMapper.hasArrivalCompleteRateData72(endDateNew, startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);
        List<RPTEveryDayCompleteEntity> unArrivalCompleteRateData72 = customerEveryDayCompleteMapper.unArrivalCompleteRateData72(endDateNew, startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);

        startDate = DateUtils.addDays(startDate, -4);
        endDateNew = DateUtils.addDays(endDate , -7);
        endQuarter = QuarterUtils.getSeasonQuarter(endDateNew);
        search.setQuarter(endQuarter);
        quarterList = QuarterUtils.getQuarters(startDate, endDate);
        List<RPTEveryDayCompleteEntity> arrivalOrderDataWeek = customerEveryDayCompleteMapper.arrivalOrderDataWeek(startDate,endDateNew,search.getAreaType(),search.getAreaId(),search.getCustomerId(),search.getQuarter());
        List<RPTEveryDayCompleteEntity> unArrivalOrderDataWeek = customerEveryDayCompleteMapper.unArrivalOrderDataWeek(startDate,endDateNew,search.getAreaType(),search.getAreaId(),search.getCustomerId(),search.getQuarter());
        //有无到货周完工单
        List<RPTEveryDayCompleteEntity> hasArrivalCompleteRateDataWeek = customerEveryDayCompleteMapper.hasArrivalCompleteRateDataWeek(endDateNew, startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);
        List<RPTEveryDayCompleteEntity> unArrivalCompleteRateDataWeek = customerEveryDayCompleteMapper.unArrivalCompleteRateDataWeek(endDateNew, startDate,endDate,search.getAreaType(),search.getAreaId(),search.getCustomerId(),quarterList);

        List<RPTEveryDayCompleteEntity> provinceList = Lists.newArrayList();
        List<RPTEveryDayCompleteEntity> cityList = Lists.newArrayList();

        try {
                Map<Long, List<RPTEveryDayCompleteEntity>> cityPlanMap = rptList.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId));
                //省完工单
                Map<Long , Integer> proCompleteMap = rptList1.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                //市完工单
                Map<Long , Integer> cityCompleteMap = rptList1.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));

                //省未完工单
                Map<Long , Integer> unCompleteOrderMap = unCompleteOrderData.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getUnCompletedOrder)));
                //市未完工单
                Map<Long , Integer> cityUnCompleteOrderMap = unCompleteOrderData.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getUnCompletedOrder)));

                //72小时前省的下单数
                Map<Long, Integer> planProMap72 = planList72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                //72小时效省完工单
                Map<Long , Integer> completeMap72 = rptList2.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getRate72)));
                //72小时效 市 完工单
                Map<Long , Integer> cityCompleteMap72 = rptList2.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getRate72)));
                //72小时前 市 的下单数
                Map<Long, Integer> planCityMap72 = planList72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));

                //48小时效 市 完工单
                Map<Long , Integer> cityCompleteMap48 = rptList3.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getArrival48)));
                //48小时前 市 的下单数
                Map<Long, Integer> planCityMap48 = planList48.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                //48小时前省下单数
                Map<Long, Integer> planProMap48 = planList48.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                //48小时效 省完工单数
                Map<Long , Integer> completeMap48 = rptList3.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getArrival48)));

                Map<Long, Integer> proArrivalDatePlanMap72 = arrivalOrderData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> cityArrivalDatePlanMap72 = arrivalOrderData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> proUnArrivalDatePlanMap72 = unArrivalOrderData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> cityUnArrivalDatePlanMap72 = unArrivalOrderData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));

                Map<Long , Integer> proArrivalDateCompleteMap72 = hasArrivalCompleteRateData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> cityArrivalDateCompleteMap72 = hasArrivalCompleteRateData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> proUnArrivalDateCompleteMap72 = unArrivalCompleteRateData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> cityUnArrivalDateCompleteMap72 = unArrivalCompleteRateData72.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));

                Map<Long, Integer> proArrivalDatePlanMapWeek = arrivalOrderDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> cityArrivalDatePlanMaWeek = arrivalOrderDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> proUnArrivalDatePlanMapWeek = unArrivalOrderDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));
                Map<Long, Integer> cityUnArrivalDatePlanMapWeek = unArrivalOrderDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getPlanOrder)));

                Map<Long , Integer> proArrivalDateCompleteMapWeek = hasArrivalCompleteRateDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> cityArrivalDateCompleteMapWeek = hasArrivalCompleteRateDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> proUnArrivalDateCompleteMapWeek = unArrivalCompleteRateDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getProvinceId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));
                Map<Long , Integer> cityUnArrivalDateCompleteMapWeek = unArrivalCompleteRateDataWeek.stream().collect(Collectors.groupingBy(RPTEveryDayCompleteEntity::getCityId, Collectors.summingInt(RPTEveryDayCompleteEntity::getCompleteOrder)));


                Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
                Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
                RPTArea province;
                RPTArea city;
                String cityName;
                Long provinceId;
                Long cityId;
                RPTEveryDayCompleteEntity rptEntity;
                for (RPTEveryDayCompleteEntity entity : proList){   //省数据开始遍历
                    rptEntity = new RPTEveryDayCompleteEntity();
                    provinceId = entity.getProvinceId();
                    province = provinceMap.get(provinceId);
                    if (province!=null){
                        rptEntity.setProvinceId(provinceId);
                        rptEntity.setProvinceName(province.getName());
                        Integer proComplete = proCompleteMap.get(entity.getProvinceId());
                        Integer planSum72 =  planProMap72.get(entity.getProvinceId());
                        Integer planSum48 = planProMap48.get(entity.getProvinceId());
                        Integer completeSum72 = completeMap72.get(entity.getProvinceId());
                        Integer completeSum48 = completeMap48.get(entity.getProvinceId());
                        Integer unCompletedSum = unCompleteOrderMap.get(entity.getProvinceId());
                        Integer proArrivalDatePlanSum72 =  proArrivalDatePlanMap72.get(entity.getProvinceId());
                        Integer proUnArrivalDatePlanSum72 = proUnArrivalDatePlanMap72.get(entity.getProvinceId());
                        Integer proArrivalDateCompleteSum72 =  proArrivalDateCompleteMap72.get(entity.getProvinceId());
                        Integer proUnArrivalDateComplete72 = proUnArrivalDateCompleteMap72.get(entity.getProvinceId());

                        Integer proArrivalDatePlanSumWeek = proArrivalDatePlanMapWeek.get(entity.getProvinceId());
                        Integer proUnArrivalDatePlanSumWeek = proUnArrivalDatePlanMapWeek.get(entity.getProvinceId());
                        Integer proArrivalDateCompleteSumWeek = proArrivalDateCompleteMapWeek.get(entity.getProvinceId());
                        Integer proUnArrivalDateCompleteSumWeek = proUnArrivalDateCompleteMapWeek.get(entity.getProvinceId());


                        rptEntity.setPlanOrder(entity.getPlanOrder());
                        rptEntity.setCompleteOrder(proComplete == null ? 0 : proComplete);
                        rptEntity.setUnCompletedOrder(unCompletedSum == null ? 0 : unCompletedSum);
                        rptEntity.setCompleteRate(new BigDecimal((rptEntity.getCompleteOrder()/rptEntity.getPlanOrder().doubleValue())*100)
                                .setScale(2, BigDecimal.ROUND_HALF_UP).toString()+"%");
                        if (planSum72 != null && planSum72 !=0) {
                            rptEntity.setCompleteRate72(new BigDecimal((completeSum72 == null ? 0 : completeSum72 / planSum72.doubleValue()) * 100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setCompleteRate72("0#F!");
                        }
                        if (planSum48!=null && planSum48 !=0) {
                            rptEntity.setArrivalCompleteRate48(new BigDecimal((completeSum48 == null ? 0 : completeSum48 / planSum48.doubleValue()) * 100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setArrivalCompleteRate48("0#F!");
                        }


                        if (proUnArrivalDatePlanSum72 !=null && proUnArrivalDatePlanSum72 != 0){
                            rptEntity.setUnArrivalDateCompletedRate72(new BigDecimal((proUnArrivalDateComplete72 == null ? 0 : proUnArrivalDateComplete72 / proUnArrivalDatePlanSum72.doubleValue())*100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setUnArrivalDateCompletedRate72("0#F!");
                        }
                        if (proArrivalDatePlanSum72 !=null && proArrivalDatePlanSum72 != 0){
                            rptEntity.setArrivalDateCompletedRate72(new BigDecimal((proArrivalDateCompleteSum72 == null ? 0 : proArrivalDateCompleteSum72 / proArrivalDatePlanSum72.doubleValue())*100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setArrivalDateCompletedRate72("0#F!");
                        }
                        if (proArrivalDatePlanSumWeek !=null && proArrivalDatePlanSumWeek != 0){
                            rptEntity.setArrivalDateCompletedRateWeek(new BigDecimal((proArrivalDateCompleteSumWeek == null ? 0 : proArrivalDateCompleteSumWeek / proArrivalDatePlanSumWeek.doubleValue())*100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setArrivalDateCompletedRateWeek("0#F!");
                        }
                        if (proUnArrivalDatePlanSumWeek !=null && proUnArrivalDatePlanSumWeek != 0){
                            rptEntity.setUnArrivalDateCompletedRateWeek(new BigDecimal((proUnArrivalDateCompleteSumWeek == null ? 0 : proUnArrivalDateCompleteSumWeek / proUnArrivalDatePlanSumWeek.doubleValue())*100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                        }else {
                            rptEntity.setUnArrivalDateCompletedRateWeek("0#F!");
                        }
                    }
                    provinceList.add(rptEntity);
                }   //省数据

                for (List<RPTEveryDayCompleteEntity> entity: cityPlanMap.values()) {    //市 数据开始遍历
                    rptEntity = new RPTEveryDayCompleteEntity();
                    provinceId = entity.get(0).getProvinceId();
                    province = provinceMap.get(provinceId);
                    cityId = entity.get(0).getCityId();
                    city = cityMap.get(cityId);
                    if (city != null) {
                        cityName = city.getName();
                        rptEntity.setProvinceId(provinceId);
                        rptEntity.setProvinceName(province.getName());
                        rptEntity.setCityId(cityId);
                        rptEntity.setCityName(cityName);
                        Integer cityCompleteSum  = cityCompleteMap.get(cityId);
                        Integer planSum72 =  planCityMap72.get(cityId);
                        Integer planSum48 = planCityMap48.get(cityId);
                        Integer completeSum72 = cityCompleteMap72.get(cityId);
                        Integer completeSum48 = cityCompleteMap48.get(cityId);
                        Integer unCompletedSum = cityUnCompleteOrderMap.get(cityId);
                        Integer cityArrivalDatePlanSum72 = cityArrivalDatePlanMap72.get(cityId);
                        Integer cityUnArrivalDatePlanSum72 = cityUnArrivalDatePlanMap72.get(cityId);
                        Integer cityArrivalDateCompleteSum72 = cityArrivalDateCompleteMap72.get(cityId);
                        Integer cityUnArrivalDateComplete72 = cityUnArrivalDateCompleteMap72.get(cityId);
                        Integer cityArrivalDatePlanSumWeek = cityArrivalDatePlanMaWeek.get(cityId);
                        Integer cityUnArrivalDatePlanSumWeek = cityUnArrivalDatePlanMapWeek.get(cityId);
                        Integer cityArrivalDateCompleteSumWeek = cityArrivalDateCompleteMapWeek.get(cityId);
                        Integer cityUnArrivalDateCompleteSumWeek = cityUnArrivalDateCompleteMapWeek.get(cityId);
                        for (RPTEveryDayCompleteEntity entityNew: entity) {
                            rptEntity.setPlanOrder(entityNew.getPlanOrder());
                            rptEntity.setCompleteOrder(cityCompleteSum == null ?0:cityCompleteSum);
                            rptEntity.setUnCompletedOrder(unCompletedSum == null?0:unCompletedSum);
                            rptEntity.setCompleteRate(new BigDecimal((rptEntity.getCompleteOrder() / rptEntity.getPlanOrder().doubleValue()) * 100)
                                    .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            if (planSum72 != null && planSum72 != 0) {
                                rptEntity.setCompleteRate72(new BigDecimal((completeSum72 == null ? 0 : completeSum72 / planSum72.doubleValue()) * 100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            } else {
                                rptEntity.setCompleteRate72("0#F!");
                            }
                            if (planSum48 != null && planSum48 != 0) {
                                rptEntity.setArrivalCompleteRate48(new BigDecimal((completeSum48 == null ? 0 : completeSum48 / planSum48.doubleValue()) * 100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            } else {
                                rptEntity.setArrivalCompleteRate48("0#F!");
                            }


                            if (cityArrivalDatePlanSum72 !=null && cityArrivalDatePlanSum72 != 0){
                                rptEntity.setArrivalDateCompletedRate72(new BigDecimal((cityArrivalDateCompleteSum72 == null ? 0 : cityArrivalDateCompleteSum72 / cityArrivalDatePlanSum72.doubleValue())*100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            }else {
                                rptEntity.setArrivalDateCompletedRate72("0#F!");
                            }
                            if (cityUnArrivalDatePlanSum72 !=null && cityUnArrivalDatePlanSum72 != 0){
                                rptEntity.setUnArrivalDateCompletedRate72(new BigDecimal((cityUnArrivalDateComplete72 == null ? 0 : cityUnArrivalDateComplete72 / cityUnArrivalDatePlanSum72.doubleValue())*100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            }else {
                                rptEntity.setUnArrivalDateCompletedRate72("0#F!");
                            }
                            if (cityArrivalDatePlanSumWeek !=null && cityArrivalDatePlanSumWeek != 0){
                                rptEntity.setArrivalDateCompletedRateWeek(new BigDecimal((cityArrivalDateCompleteSumWeek == null ? 0 : cityArrivalDateCompleteSumWeek /cityArrivalDatePlanSumWeek.doubleValue())*100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            }else {
                                rptEntity.setArrivalDateCompletedRateWeek("0#F!");
                            }
                            if (cityUnArrivalDatePlanSumWeek !=null && cityUnArrivalDatePlanSumWeek != 0){
                                rptEntity.setUnArrivalDateCompletedRateWeek(new BigDecimal((cityUnArrivalDateCompleteSumWeek == null ? 0 : cityUnArrivalDateCompleteSumWeek / cityUnArrivalDatePlanSumWeek.doubleValue())*100)
                                        .setScale(2, BigDecimal.ROUND_HALF_UP).toString() + "%");
                            }else {
                                rptEntity.setUnArrivalDateCompletedRateWeek("0#F!");
                            }

                            cityList.add(rptEntity);
                        }
                    }
                }
            cityList = cityList.stream().sorted(Comparator.comparing(RPTEveryDayCompleteEntity::getProvinceId)
                    .thenComparing(RPTEveryDayCompleteEntity :: getCityId)
                    .thenComparing(RPTEveryDayCompleteEntity :: getPlanOrder)).collect(Collectors.toList());
        }catch (Exception e) {
            log.error("每日完工时效数据异常{}",e.getMessage());
        }
        map.put(RPTEveryDayCompleteEntity.MAP_KEY_PROVINCELIST, provinceList);
        map.put(RPTEveryDayCompleteEntity.MAP_KEY_CITYLIST, cityList);
        return map;
    }



    private List<String> getQuarters(Date endDate) {
        endDate = DateUtils.getEndOfDay(endDate);
        Date goLiveDate = RptCommonUtils.getGoLiveDate();
        List<String> quarters = QuarterUtils.getQuarters(goLiveDate, endDate);
        int size = quarters.size();
        if (size > 3) {
            quarters = quarters.subList(size-3, size);
        }
        return quarters;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTEveryDayCompleteSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTEveryDayCompleteSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getStartDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = customerEveryDayCompleteMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    /**
     *
     *每日完工时效导出
     * @return
     */
    public SXSSFWorkbook areaOrderCompleteRateRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTEveryDayCompleteSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTEveryDayCompleteSearch.class);
            String day = DateUtils.getDay(DateUtils.timeStampToDate(searchCondition.getStartDate()));
            Map<String, List<RPTEveryDayCompleteEntity>> entityMap = getEveryDayComplete(searchCondition);

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 11));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER,day);
            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            ExportExcel.createCell(headSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单");

            ExportExcel.createCell(headSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成单");

            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成率");

            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "72H完成率");

            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货48H完成率");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "无到货72H完成率");
            ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "无到货周完成率");
            ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货72H完成率");
            ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货周完成率");

            ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工单");

            List<RPTEveryDayCompleteEntity> pList = entityMap.get(RPTEveryDayCompleteEntity.MAP_KEY_PROVINCELIST);
            List<RPTEveryDayCompleteEntity> cList = entityMap.get(RPTEveryDayCompleteEntity.MAP_KEY_CITYLIST);

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            if (pList != null && pList.size() > 0) {
                int pCount = pList.size();
                // 循环读取所有的省
                for (int i = 0; i < pCount; i++) {
                    RPTEveryDayCompleteEntity province = pList.get(i);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int pColumnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == province.getProvinceName() ? "" : province.getProvinceName());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getPlanOrder());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getCompleteOrder());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getCompleteRate()==""?"0.00%":province.getCompleteRate());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getCompleteRate72()==""?"0.00%":province.getCompleteRate72());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getArrivalCompleteRate48()==""?"0.00%":province.getArrivalCompleteRate48());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getUnArrivalDateCompletedRate72());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getUnArrivalDateCompletedRateWeek());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getArrivalDateCompletedRate72());
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getArrivalDateCompletedRateWeek());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getUnCompletedOrder());

                    //循环读取省下的市
                    for (RPTEveryDayCompleteEntity city : cList) {
                        if (city.getProvinceId().equals(province.getProvinceId())) {

                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            int cColumnIndex = 0;

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                            dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getPlanOrder());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCompleteOrder());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCompleteRate()==""?"0.00%":city.getCompleteRate());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCompleteRate72()==""?"0.00%":city.getCompleteRate72());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getArrivalCompleteRate48()==""?"0.00%":city.getArrivalCompleteRate48());

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getUnArrivalDateCompletedRate72());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getUnArrivalDateCompletedRateWeek());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getArrivalDateCompletedRate72());
                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getArrivalDateCompletedRateWeek());

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getUnCompletedOrder());

                        }

                    }// 循环读取省下的市

                }// 循环读取所有的省

            }
        } catch (Exception e) {
            log.error("【EveryDayCompleteService.areaOrderCompleteRateRptExport】省市每日完工时效报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}
