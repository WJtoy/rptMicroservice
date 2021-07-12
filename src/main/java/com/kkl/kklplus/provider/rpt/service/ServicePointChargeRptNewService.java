package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.RPTServicePointChargeEntity;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointChargeNewRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointChargeRptNewService extends RptBaseService {

    @Resource
    private ServicePointChargeNewRptMapper servicePointChargeNewRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;


    private List<RPTServicePointChargeEntity> getServicePointChargeFromWeb(Date queryDate) {
        List<RPTServicePointChargeEntity> result = Lists.newArrayList();
        //计算查询条件
        final int pageSize = 20000;
        final String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        final Date now = new Date();
        final Date beginDate = DateUtils.getStartDayOfMonth(DateUtils.getStartOfDay(queryDate));
        Date tempEndDate = DateUtils.addMonth(beginDate, 1);
        if (DateUtils.getYearMonth(queryDate).equals(DateUtils.getYearMonth(now))) {
            tempEndDate = DateUtils.getStartOfDay(now);
        }
        final Date endDate = tempEndDate;
        final int selectedYear = DateUtils.getYear(beginDate);
        final int selectedMonth = DateUtils.getMonth(beginDate);
        final int yearMonth = StringUtils.toInteger(DateUtils.getYearMonth(beginDate));
        final Date preMonthStartDay = DateUtils.addMonth(beginDate, -1);
        final int preSelectedYear = DateUtils.getYear(preMonthStartDay);
        final int preSelectedMonth = DateUtils.getMonth(preMonthStartDay);

        Date preBeginDate = DateUtils.addMonth(beginDate,-1);
        Date preEndDate = DateUtils.addMonth(endDate,-1);
        String preQuarter = QuarterUtils.getSeasonQuarter(beginDate);

        //region 分页读取Web原始数据
        List<RPTServicePointChargeEntity> payableAList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getPayableAList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> payableBList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getPayableBList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> paidAmountList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getPaidAmountList(selectedYear, selectedMonth, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> lastBalanceList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getLastMonthBalanceList(preSelectedYear, preSelectedMonth, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> theBalanceList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getTheBalanceList(selectedYear, selectedMonth, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> completedChargeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getCompletedChargeList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> returnChargeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getReturnChargeList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> timelineUrgentInsurancePraiseTaxInfoList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getTimelineUrgentInsurancePraiseTaxInfoList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> platformFeeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getPlatformFeeList(quarter, beginDate, endDate, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> engineerDepositFeeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getAllEngineerDepositFees(quarter, beginDate, endDate, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> rechargeDepositFeeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getAllRechargeDepositFees(quarter, beginDate, endDate, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> preEngineerDepositFeeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getAllEngineerDepositFees(preQuarter, preBeginDate, preEndDate, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> preRechargeDepositFeeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getAllRechargeDepositFees(preQuarter, preBeginDate, preEndDate, startLimit, pageSize);
            }
        }, pageSize);

        List<RPTServicePointChargeEntity> otherTravelChargeList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getOtherTravelChargeList(beginDate, endDate, quarter, startLimit, pageSize);
            }
        }, pageSize);
        List<RPTServicePointChargeEntity> completeQtyList = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getCompleteQtyList(beginDate, endDate, startLimit, pageSize);
            }
        }, pageSize);
        //endregion 分页读取Web原始数据

        Map<String, RPTServicePointChargeEntity> payableAMap = payableAList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> payableBMap = payableBList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> paidAmountMap = paidAmountList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> lastBalanceMap = lastBalanceList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> theBalanceListMap = theBalanceList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> completedChargeMap = completedChargeList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> returnChargeMap = returnChargeList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> timelineUrgentInsurancePraiseTaxInfoMap = timelineUrgentInsurancePraiseTaxInfoList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<Long, RPTServicePointChargeEntity> platformFeeMap = platformFeeList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Map<Long, RPTServicePointChargeEntity> engineerDepositFeeMap = engineerDepositFeeList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Map<Long, RPTServicePointChargeEntity> rechargeDepositFeeMap = rechargeDepositFeeList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Map<Long, RPTServicePointChargeEntity> preEngineerDepositFeeMap = preEngineerDepositFeeList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Map<Long, RPTServicePointChargeEntity> preRechargeDepositFeeMap = preRechargeDepositFeeList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Map<String, RPTServicePointChargeEntity> otherTravelChargeMap = otherTravelChargeList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        Map<String, RPTServicePointChargeEntity> completeQtyMap = completeQtyList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        Set<String> keySet = Sets.newHashSet();
        keySet.addAll(payableAMap.keySet());
        keySet.addAll(payableBMap.keySet());
        keySet.addAll(paidAmountMap.keySet());
        keySet.addAll(lastBalanceMap.keySet());
        keySet.addAll(theBalanceListMap.keySet());
        keySet.addAll(completedChargeMap.keySet());
        keySet.addAll(returnChargeMap.keySet());
        keySet.addAll(timelineUrgentInsurancePraiseTaxInfoMap.keySet());
        keySet.addAll(otherTravelChargeMap.keySet());
        keySet.addAll(completeQtyMap.keySet());

        Set<Long> spIdSet = Sets.newHashSet();
        spIdSet.addAll(payableAList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(payableBList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(paidAmountList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(lastBalanceList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(theBalanceList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(completedChargeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(returnChargeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(timelineUrgentInsurancePraiseTaxInfoList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(platformFeeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(engineerDepositFeeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(rechargeDepositFeeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(preEngineerDepositFeeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(preRechargeDepositFeeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(otherTravelChargeList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        spIdSet.addAll(completeQtyList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        List<Long> servicePointIds = Lists.newArrayList(spIdSet);

        Map<Long, RPTArea> areaMap = MSAreaUtils.getAreaMap(RPTArea.TYPE_VALUE_COUNTY);
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Lists.newArrayList("id", "primaryId", "paymentType", "areaId"), null);
        List<Long> engineerIds = servicePointMap.values().stream().map(MDServicePointViewModel::getPrimaryId).distinct().collect(Collectors.toList());
        Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
        if (!engineerIds.isEmpty()) {
            List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Lists.newArrayList("id", "name", "appFlag"));
            if (engineerList != null && !engineerList.isEmpty()) {
                engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
            }
        }

        List<RPTServicePointChargeEntity> list = Lists.newArrayListWithCapacity(keySet.size());
        RPTServicePointChargeEntity newObj, tempObj;
        MDServicePointViewModel servicePoint;
        RPTEngineer engineer;
        RPTArea area;
        for (String key : keySet) {
            newObj = new RPTServicePointChargeEntity();
            newObj.setQuarter(quarter);
            newObj.setYearMonth(yearMonth);
            String[] idsStr = StringUtils.split(key, ":");
            newObj.setServicePointId(StringUtils.toLong(idsStr[0]));
            newObj.setProductCategoryId(StringUtils.toLong(idsStr[1]));

            servicePoint = servicePointMap.get(newObj.getServicePointId());
            if (servicePoint != null) {
                newObj.setPaymentType(StringUtils.toInteger(servicePoint.getPaymentType()));
                area = areaMap.get(servicePoint.getAreaId());
                if (area != null) {
                    String[] split = area.getParentIds().split(",");
                    if (split.length == 4) {
                        newObj.setCountyId(area.getId());
                        newObj.setCityId(StringUtils.toLong(split[3]));
                        newObj.setProvinceId(StringUtils.toLong(split[2]));
                    }
                }
                engineer = engineerMap.get(servicePoint.getPrimaryId());
                if (engineer != null) {
                    newObj.setAppFlag(StringUtils.toInteger(engineer.getAppFlag()));
                }
            }

            tempObj = payableAMap.get(key);
            if (tempObj != null) {
                newObj.setPayableA(StringUtils.toDouble(tempObj.getPayableA()));
            }
            tempObj = payableBMap.get(key);
            if (tempObj != null) {
                newObj.setPayableB(StringUtils.toDouble(tempObj.getPayableB()));
            }
            tempObj = paidAmountMap.get(key);
            if (tempObj != null) {
                newObj.setPaidAmount(StringUtils.toDouble(tempObj.getPaidAmount()));
            }
            tempObj = lastBalanceMap.get(key);
            if (tempObj != null) {
                newObj.setLastMonthBalance(StringUtils.toDouble(tempObj.getLastMonthBalance()));
            }
            tempObj = theBalanceListMap.get(key);
            if (tempObj != null) {
                newObj.setTheBalance(StringUtils.toDouble(tempObj.getTheBalance()));
            }
            tempObj = completedChargeMap.get(key);
            if (tempObj != null) {
                newObj.setCompletedCharge(StringUtils.toDouble(tempObj.getCompletedCharge()));
            }
            tempObj = returnChargeMap.get(key);
            if (tempObj != null) {
                newObj.setReturnCharge(StringUtils.toDouble(tempObj.getReturnCharge()));
            }
            tempObj = timelineUrgentInsurancePraiseTaxInfoMap.get(key);
            if (tempObj != null) {
                newObj.setTimelinessCharge(StringUtils.toDouble(tempObj.getTimelinessCharge()));
                newObj.setCustomerTimelinessCharge(StringUtils.toDouble(tempObj.getCustomerTimelinessCharge()));
                newObj.setUrgentCharge(StringUtils.toDouble(tempObj.getUrgentCharge()));
                newObj.setInsuranceCharge(StringUtils.toDouble(tempObj.getInsuranceCharge()));
                newObj.setPraiseFee(StringUtils.toDouble(tempObj.getPraiseFee()));
                newObj.setTaxFee(StringUtils.toDouble(tempObj.getTaxFee()));
                newObj.setInfoFee(StringUtils.toDouble(tempObj.getInfoFee()));
            }
            tempObj = otherTravelChargeMap.get(key);
            if (tempObj != null) {
                newObj.setEngineerOtherCharge(StringUtils.toDouble(tempObj.getEngineerOtherCharge()));
                newObj.setEngineerTravelCharge(StringUtils.toDouble(tempObj.getEngineerTravelCharge()));
            }
            tempObj = completeQtyMap.get(key);
            if (tempObj != null) {
                newObj.setCompleteQty(StringUtils.toInteger(tempObj.getCompleteQty()));
            }
            newObj.setPayableAmount(newObj.getPayableA() + newObj.getPayableB());
            list.add(newObj);
        }

        Map<Long, List<RPTServicePointChargeEntity>> objMap = list.stream().collect(Collectors.groupingBy(RPTServicePointChargeEntity::getServicePointId));
        List<RPTServicePointChargeEntity> summaryList = Lists.newArrayListWithCapacity(objMap.size());
        RPTServicePointChargeEntity summaryObj;
        double platformFee = 0;
        double engineerDepositFee = 0;
        double rechargeDepositFee = 0;
        double preEngineerDepositFee = 0;
        double preRechargeDepositFee = 0;
        for (List<RPTServicePointChargeEntity> subList : objMap.values()) {
            if (!subList.isEmpty()) {
                tempObj = platformFeeMap.get(subList.get(0).getServicePointId());
                platformFee = tempObj == null ? 0 : StringUtils.toDouble(tempObj.getPlatformFee());
                tempObj = engineerDepositFeeMap.get(subList.get(0).getServicePointId());
                engineerDepositFee = tempObj == null ? 0 : StringUtils.toDouble(tempObj.getEngineerDeposit());
                tempObj = rechargeDepositFeeMap.get(subList.get(0).getServicePointId());
                rechargeDepositFee = tempObj == null ? 0 : StringUtils.toDouble(tempObj.getRechargeDeposit());
                tempObj = preEngineerDepositFeeMap.get(subList.get(0).getServicePointId());
                preEngineerDepositFee = tempObj == null ? 0 : StringUtils.toDouble(tempObj.getEngineerDeposit());
                tempObj = preRechargeDepositFeeMap.get(subList.get(0).getServicePointId());
                preRechargeDepositFee = tempObj == null ? 0 : StringUtils.toDouble(tempObj.getRechargeDeposit());
                summaryObj = createSummaryEntity(subList, platformFee,engineerDepositFee,rechargeDepositFee,preEngineerDepositFee,preRechargeDepositFee);
                if (platformFee != 0) {
                    setPlatformFee(subList, platformFee);
                }
                summaryList.add(summaryObj);
            }
        }

        result.addAll(summaryList);
        result.addAll(list.stream().filter(i -> i.getDelFlag() == RPTServicePointChargeEntity.DEL_FLAG_NORMAL).collect(Collectors.toList()));
        return result;
    }

    private RPTServicePointChargeEntity createSummaryEntity(List<RPTServicePointChargeEntity> list, double platformFee,double engineerDepositFee,double rechargeDepositFee,double preEngineerDepositFee,double preRechargeDepositFee) {
        RPTServicePointChargeEntity firstObj = list.get(0);
        RPTServicePointChargeEntity newObj = new RPTServicePointChargeEntity();
        newObj.setQuarter(firstObj.getQuarter());
        newObj.setYearMonth(firstObj.getYearMonth());
        newObj.setServicePointId(firstObj.getServicePointId());
        newObj.setPaymentType(firstObj.getPaymentType());
        newObj.setCountyId(firstObj.getCountyId());
        newObj.setCityId(firstObj.getCityId());
        newObj.setProvinceId(firstObj.getProvinceId());
        newObj.setAppFlag(firstObj.getAppFlag());
        newObj.setProductCategoryId(0L);
        newObj.setPlatformFee(Math.abs(platformFee));
        newObj.setEngineerDeposit(0 - engineerDepositFee);//以负数展现
        newObj.setRechargeDeposit(0 - rechargeDepositFee);
        newObj.setPreEngineerDeposit(0 - preEngineerDepositFee);
        newObj.setPreRechargeDeposit(0 - preRechargeDepositFee);

        for (RPTServicePointChargeEntity item : list) {
            if (item.getProductCategoryId() == 0) {
                item.setDelFlag(RPTServicePointChargeEntity.DEL_FLAG_DELETE);
            }
            newObj.setPayableAmount(newObj.getPayableAmount() + item.getPayableAmount());
            newObj.setPaidAmount(newObj.getPaidAmount() + item.getPaidAmount());
            newObj.setLastMonthBalance(newObj.getLastMonthBalance() + item.getLastMonthBalance());
            newObj.setTheBalance(newObj.getTheBalance() + item.getTheBalance());
            newObj.setCompletedCharge(newObj.getCompletedCharge() + item.getCompletedCharge());
            newObj.setReturnCharge(newObj.getReturnCharge() + item.getReturnCharge());
            newObj.setTimelinessCharge(newObj.getTimelinessCharge() + item.getTimelinessCharge());
            newObj.setCustomerTimelinessCharge(newObj.getCustomerTimelinessCharge() + item.getCustomerTimelinessCharge());
            newObj.setUrgentCharge(newObj.getUrgentCharge() + item.getUrgentCharge());
            newObj.setInsuranceCharge(newObj.getInsuranceCharge() + item.getInsuranceCharge());
            newObj.setPraiseFee(newObj.getPraiseFee() + item.getPraiseFee());
            newObj.setTaxFee(newObj.getTaxFee() + item.getTaxFee());
            newObj.setInfoFee(newObj.getInfoFee() + item.getInfoFee());
            newObj.setEngineerOtherCharge(newObj.getEngineerOtherCharge() + item.getEngineerOtherCharge());
            newObj.setEngineerTravelCharge(newObj.getEngineerTravelCharge() + item.getEngineerTravelCharge());
            newObj.setCompleteQty(newObj.getCompleteQty() + item.getCompleteQty());
        }

        return newObj;
    }

    private void setPlatformFee(List<RPTServicePointChargeEntity> list, double platformFee) {
        double totalPlatformFee = Math.abs(platformFee);
        double fee;
        RPTServicePointChargeEntity item;
        for (int i = 0; i < list.size(); i++) {
            item = list.get(i);
            if (i == list.size() - 1) {
                item.setPlatformFee(totalPlatformFee);
            } else {
                fee = CurrencyUtil.round2(item.getPaidAmount() * RptCommonUtils.getPlatformFeeRate());
                if (fee >= totalPlatformFee) {
                    item.setPlatformFee(totalPlatformFee);
                    totalPlatformFee = 0;
                } else {
                    if (totalPlatformFee > 0) {
                        item.setPlatformFee(fee);
                        totalPlatformFee = totalPlatformFee - fee;
                    }
                }
            }
            item.setPlatformFee(Math.abs(item.getPlatformFee()));
        }
    }

    private void insertServicePointCharge(RPTServicePointChargeEntity entity, String quarter, int systemId) {
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            servicePointChargeNewRptMapper.insert(entity);
        } catch (Exception e) {
            log.error("【ServicePointChargeRptNewService.insertServicePointCharge】sId: {}, pId: {}, yearmonth: {}, errorMsg: {}",
                    entity.getServicePointId(), entity.getProductCategoryId(), entity.getYearMonth(),
                    Exceptions.getStackTraceAsString(e));
        }
    }

    private void updateServicePointCharge(RPTServicePointChargeEntity entity, String quarter, int systemId, Long id) {
        try {
            entity.setId(id);
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            servicePointChargeNewRptMapper.update(entity);
        } catch (Exception e) {
            log.error("【ServicePointChargeRptNewService.updateServicePointCharge】sId: {}, pId: {}, yearmonth: {}, errorMsg: {}",
                    entity.getServicePointId(), entity.getProductCategoryId(), entity.getYearMonth(),
                    Exceptions.getStackTraceAsString(e));
        }
    }

    public void saveServicePointChargeToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointChargeEntity> list = getServicePointChargeFromWeb(date);
            if (!list.isEmpty()) {
                int systemId = RptCommonUtils.getSystemId();
                String quarter = QuarterUtils.getSeasonQuarter(date);
                for (RPTServicePointChargeEntity item : list) {
                    insertServicePointCharge(item, quarter, systemId);
                }
                int yearMonth = StringUtils.toInteger(DateUtils.getYearMonth(date));
                log.error("【重建中间表日志】【ServicePointChargeRptNewService.saveServicePointChargeToRptDB】quarter: {}, systemId: {}, yearmonth: {}", quarter, systemId, yearMonth);
            }
        }
    }

    /**
     * 删除中间表某一个月的数据
     */
    private void deleteServicePointChargeFromRptDB(Date date) {
        if (date != null) {
            int systemId = RptCommonUtils.getSystemId();
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int yearMonth = StringUtils.toInteger(DateUtils.getYearMonth(date));
            servicePointChargeNewRptMapper.delete(quarter, systemId, yearMonth);
            log.error("【重建中间表日志】【ServicePointChargeRptNewService.deleteServicePointChargeFromRptDB】quarter: {}, systemId: {}, yearmonth: {}", quarter, systemId, yearMonth);
        }
    }

    /**
     * 从中间表读取"网点 + 品类"对应的主键id
     */
    public Map<String, Long> getServicePointChargeIdsFromRptDB(String quarter, int systemId, int yearmonth) {
        Map<String, Long> result = Maps.newHashMap();
        List<RPTServicePointChargeEntity> list = PageSearchUtils.exec(new PageSearchUtils.QueryAction<RPTServicePointChargeEntity>() {
            @Override
            public List<RPTServicePointChargeEntity> query(int startLimit, int pageSize) {
                return servicePointChargeNewRptMapper.getServicePointChargeIds(quarter, systemId, yearmonth, startLimit, pageSize);
            }
        }, 20000);
        for (RPTServicePointChargeEntity item : list) {
            result.put(item.getServicePointId() + ":" + item.getProductCategoryId(), item.getId());
        }
        return result;
    }

    public void updateServicePointChargeToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointChargeEntity> list = getServicePointChargeFromWeb(date);
            if (!list.isEmpty()) {
                int systemId = RptCommonUtils.getSystemId();
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int yearMonth = StringUtils.toInteger(DateUtils.getYearMonth(date));
                Map<String, Long> servicePointChargeIdMap = getServicePointChargeIdsFromRptDB(quarter, systemId, yearMonth);
                Long servicePointChargeId;
                for (RPTServicePointChargeEntity item : list) {
                    servicePointChargeId = servicePointChargeIdMap.get(item.getServicePointId() + ":" + item.getProductCategoryId());
                    if (servicePointChargeId != null) {
                        updateServicePointCharge(item, quarter, systemId, servicePointChargeId);
                    } else {
                        insertServicePointCharge(item, quarter, systemId);
                    }
                }
                log.error("【重建中间表日志】【ServicePointChargeRptNewService.updateServicePointChargeToRptDB】quarter: {}, systemId: {}, yearmonth: {}", quarter, systemId, yearMonth);
            }
        }
    }

    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() <= endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveServicePointChargeToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
//                            updateServicePointChargeToRptDB(beginDate);
                            deleteServicePointChargeFromRptDB(beginDate);
                            saveServicePointChargeToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointChargeFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointChargeRptNewService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


}
