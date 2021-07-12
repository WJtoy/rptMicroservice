package com.kkl.kklplus.provider.rpt.service;


import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;

import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.*;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.*;
import com.kkl.kklplus.provider.rpt.mapper.GradedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.provider.rpt.utils.web.RPTOrderItemUtils;
import com.kkl.kklplus.starter.redis.config.RedisGsonService;
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

import static com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW;
import static com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity.RPT_ROW_NUMBER_SUMROW;
import static com.kkl.kklplus.provider.rpt.service.RptBaseService.*;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class GradedOrderRptService {

    @Resource
    private GradedOrderRptMapper gradedOrderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private RedisGsonService redisGsonService;

    @Autowired
    private AreaCacheService areaCacheService;

    @Autowired
    private KeFuUtils keFuUtils;

    /**
     * 根据时间获取原表中的数据
     * @param quaryDate
     * @return
     */
    public List<RPTGradedOrderEntity> getGradedOrderData(Date quaryDate){
        Date startDate = DateUtils.getDateStart(quaryDate);
        Date endDate = DateUtils.getDateEnd(quaryDate);
        List<RPTGradedOrderEntity> gradedOrderData = gradedOrderRptMapper.getGradedOrderData(startDate, endDate);
        List<RPTGradedOrderEntity> crushOrderData = gradedOrderRptMapper.getCrushOrderData(startDate, endDate);
        List<TwoTuple<Long,Long>> gradeCreateByList = gradedOrderRptMapper.getGradeCreateByList(startDate, endDate);
        Map<Long, List<Long>> gradeCreateByMap = gradeCreateByList.stream().collect(Collectors.groupingBy(TwoTuple::getAElement, Collectors.mapping(TwoTuple::getBElement, Collectors.toList())));

        Map<Long, RPTGradedOrderEntity> crushMap = new HashMap<>();
        for (RPTGradedOrderEntity entity:crushOrderData) {
            crushMap.put(entity.getOrderId(),entity);
        }
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap =areaCacheService.getAllCityMap();
        //获取所有的客服
        Map<Long, RPTUser> kefuMap = MSUserUtils.getMapByUserType(RPTUser.USER_TYPE_KEFU);
        Map<Long, RPTProductCategory> allProductCategoryMap = MDUtils.getAllProductCategoryMap();
        List<Long> servicePointIds = gradedOrderData.stream().map(RPTGradedOrderEntity::getServicePointId).collect(Collectors.toList());
        List<Long> customerIds = gradedOrderData.stream().map(RPTGradedOrderEntity::getCustomerId).collect(Collectors.toList());
        List<Long> distinctServicePointIds = servicePointIds.stream().distinct().collect(Collectors.toList());
        Set<Long> distinctCustomerIds = customerIds.stream().distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(distinctCustomerIds), Arrays.asList(fieldsArray));
        Map<Long, RPTServicePoint> servicePointMap = msServicePointService.getServicePointMap(distinctServicePointIds);
        List<RPTOrderItem> orderItemList = new ArrayList<>();
        for (RPTGradedOrderEntity entity : gradedOrderData) {
            //从微服务中获取网点信息
            if (servicePointMap!=null){
                RPTServicePoint rptServicePoint = servicePointMap.get(entity.getServicePointId());
                if (rptServicePoint!=null){
                    entity.setServicePointName(rptServicePoint.getName());
                    entity.setServicePointNo(rptServicePoint.getServicePointNo());
                }else {
                    entity.setServicePointName("");
                    entity.setServicePointNo("");
                }
            }
            //设置客评人
            if (gradeCreateByMap!=null){
                List<Long> gradeCreateBy = gradeCreateByMap.get(entity.getOrderId());
                if (gradeCreateBy!=null && gradeCreateBy.size()>0){
                    RPTUser rptUser = kefuMap.get(gradeCreateBy.get(0));
                    if (rptUser!=null){
                        entity.setGradeBy(gradeCreateBy.get(0));
                    }else {
                        entity.setGradeBy(entity.getKefuId());
                    }
                }else {
                    entity.setGradeBy(entity.getKefuId());
                }
            }
//            if (entity.getGradeBy()!=null){
//                RPTUser rptUser = kefuMap.get(entity.getGradeBy());
//                if (rptUser==null){
//                    entity.setGradeBy(entity.getKefuId());
//                }
//            }else {
//                entity.setGradeBy(entity.getKefuId());
//            }
            //从微服务中获取客户信息
            if (customerMap!=null) {
                RPTCustomer rptCustomer = customerMap.get(entity.getCustomerId());
                if (rptCustomer != null) {
                    entity.setCustomerName(rptCustomer.getName());
                }else {
                    entity.setCustomerName("");
                }
            }
            entity.setCreateDate(entity.getCreateDateD().getTime());
            if (entity.getFirstPlanDateD()!=null) {
                entity.setFirstPlanDate(entity.getFirstPlanDateD().getTime());
            }
            if (entity.getPlanDateD()!=null) {
                entity.setPlanDate(entity.getPlanDateD().getTime());
            }
            if (entity.getArrivalDate()!=null) {
                entity.setArrivalDt(entity.getArrivalDate().getTime());
            }
            if (entity.getAppCompleteDate()!=null) {
                entity.setAppCompleteDt(entity.getAppCompleteDate().getTime());
            }else {
                entity.setAppCompleteDt(0l);
            }
            entity.setCloseDate(entity.getCloseDateD().getTime());
            //获取服务品类的名称
            if (allProductCategoryMap!=null){
                RPTProductCategory rptProductCategory = allProductCategoryMap.get(entity.getProductCategoryId());
                if (rptProductCategory!=null){
                    entity.setProductCategoryName(rptProductCategory.getName());
                }else {
                    entity.setProductCategoryName("");
                }
            }else {
                entity.setProductCategoryName("");
            }
            //获取突击单信息
            RPTGradedOrderEntity crushEntity = crushMap.get(entity.getOrderId());
            if (crushEntity!=null){
                entity.setCrushFlag(1);
                entity.setCrushCreateBy(crushEntity.getCrushCreateBy());
                entity.setCrushCreateDate(crushEntity.getCrushCreateDateD().getTime());

                long crushHour = (entity.getCloseDate() - entity.getCrushCreateDate()) / (1000 * 60 * 60);
                if (crushHour < 96 ){
                    entity.setCrush96HourComplete(1);
                }
            }
//            entity.setOrderItems(RPTOrderItemUtils.fromOrderItemsJson(entity.getOrderItemJson()));
            entity.setOrderItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(entity.getOrderItemPb()));
            orderItemList.addAll(entity.getOrderItems());
            //获取区域名
            RPTArea rptArea = areaMap.get(entity.getCountyId());
            if (rptArea!=null) {
                RPTArea city = rptArea.getParent();
                if (city!=null){
                    entity.setCityId(city.getId());
                    RPTArea cityArea = cityMap.get(city.getId());
                    if (cityArea!=null){
                        RPTArea parent = cityArea.getParent();
                        if (parent!=null){
                            entity.setProvinceId(parent.getId());
                        }
                    }
                }
            }
        }
        RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
        return gradedOrderData;
    }

    /**
     * 根据消息队列获取客评工单信息
     * @param quarter
     * @param orderId
     * @return
     */
    public RPTGradedOrderEntity getGradedOrderOfMQ(String quarter,Long orderId){
        RPTGradedOrderEntity gradedOrderDataOfMQ = gradedOrderRptMapper.getGradedOrderDataOfMQ(quarter, orderId);
        RPTGradedOrderEntity crushOrderDataOfMQ = gradedOrderRptMapper.getCrushOrderDataOfMQ(quarter, orderId);
        List<TwoTuple<Long,Long>> gradeCreateByList = gradedOrderRptMapper.getGradeCreateByOfMQ(quarter, orderId);
        Map<Long, List<Long>> gradeCreateByMap = gradeCreateByList.stream().collect(Collectors.groupingBy(TwoTuple::getAElement, Collectors.mapping(TwoTuple::getBElement, Collectors.toList())));

        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        //获取所有的客服
        Map<Long, RPTUser> kefuMap = MSUserUtils.getMapByUserType(RPTUser.USER_TYPE_KEFU);
        Map<Long, RPTProductCategory> allProductCategoryMap = MDUtils.getAllProductCategoryMap();
        List<RPTOrderItem> orderItemList = new ArrayList<>();
        List<Long> servicePointIds  = Lists.newArrayList();
        Set<Long> customerIds = new HashSet<>();
        if (gradedOrderDataOfMQ!=null && gradedOrderDataOfMQ.getServicePointId()!=null){
            servicePointIds.add(gradedOrderDataOfMQ.getServicePointId());
        }
        if (gradedOrderDataOfMQ!= null && gradedOrderDataOfMQ.getCustomerId()!=null){
            customerIds.add(gradedOrderDataOfMQ.getCustomerId());
        }
        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Map<Long, RPTServicePoint> servicePointMap = msServicePointService.getServicePointMap(servicePointIds);
        if (servicePointMap!=null){
            RPTServicePoint rptServicePoint = servicePointMap.get(gradedOrderDataOfMQ.getServicePointId());
            if (rptServicePoint!=null){
                gradedOrderDataOfMQ.setServicePointName(rptServicePoint.getName());
                gradedOrderDataOfMQ.setServicePointNo(rptServicePoint.getServicePointNo());
            }else {
                gradedOrderDataOfMQ.setServicePointName("");
                gradedOrderDataOfMQ.setServicePointNo("");
            }
        }

        //设置客评人
//        if (gradedOrderDataOfMQ.getGradeBy()!=null){
//            RPTUser rptUser = kefuMap.get(gradedOrderDataOfMQ.getGradeBy());
//            if (rptUser==null){
//                gradedOrderDataOfMQ.setGradeBy(gradedOrderDataOfMQ.getKefuId());
//            }
//        }else {
//            gradedOrderDataOfMQ.setGradeBy(gradedOrderDataOfMQ.getKefuId());
//        }
        gradedOrderDataOfMQ.setCloseDate(gradedOrderDataOfMQ.getCloseDateD().getTime());
        //获取突击单信息
        if (crushOrderDataOfMQ!=null){
            gradedOrderDataOfMQ.setCrushFlag(1);
            gradedOrderDataOfMQ.setCrushCreateBy(crushOrderDataOfMQ.getCrushCreateBy());
            gradedOrderDataOfMQ.setCrushCreateDate(crushOrderDataOfMQ.getCrushCreateDateD().getTime());
            long crushHour = (gradedOrderDataOfMQ.getCloseDate() - gradedOrderDataOfMQ.getCrushCreateDate()) / (1000 * 60 * 60);
            if (crushHour < 96 ){
                gradedOrderDataOfMQ.setCrush96HourComplete(1);
            }
        }
        //设置客评人
        if (gradeCreateByMap!=null){
            List<Long> gradeCreateBy = gradeCreateByMap.get(gradedOrderDataOfMQ.getOrderId());
            if (gradeCreateBy!=null && gradeCreateBy.size()>0){
                RPTUser rptUser = kefuMap.get(gradeCreateBy.get(0));
                if (rptUser!=null){
                    gradedOrderDataOfMQ.setGradeBy(gradeCreateBy.get(0));
                }else {
                    gradedOrderDataOfMQ.setGradeBy(gradedOrderDataOfMQ.getKefuId());
                }
            }else {
                gradedOrderDataOfMQ.setGradeBy(gradedOrderDataOfMQ.getKefuId());
            }
        }
        //从微服务中获取客户信息
        if (customerMap!=null) {
            RPTCustomer rptCustomer = customerMap.get(gradedOrderDataOfMQ.getCustomerId());
            if (rptCustomer != null) {
                gradedOrderDataOfMQ.setCustomerName(rptCustomer.getName());
            }else {
                gradedOrderDataOfMQ.setCustomerName("");
            }
        }
        gradedOrderDataOfMQ.setCreateDate(gradedOrderDataOfMQ.getCreateDateD().getTime());
        if (gradedOrderDataOfMQ.getFirstPlanDateD()!=null) {
            gradedOrderDataOfMQ.setFirstPlanDate(gradedOrderDataOfMQ.getFirstPlanDateD().getTime());
        }
        if (gradedOrderDataOfMQ.getPlanDateD()!=null) {
            gradedOrderDataOfMQ.setPlanDate(gradedOrderDataOfMQ.getPlanDateD().getTime());
        }
        if (gradedOrderDataOfMQ.getArrivalDate()!=null) {
            gradedOrderDataOfMQ.setArrivalDt(gradedOrderDataOfMQ.getArrivalDate().getTime());
        }
        if (gradedOrderDataOfMQ.getAppCompleteDate()!=null) {
            gradedOrderDataOfMQ.setAppCompleteDt(gradedOrderDataOfMQ.getAppCompleteDate().getTime());
        }else {
            gradedOrderDataOfMQ.setAppCompleteDt(0L);
        }

        //获取服务品类的名称
        if (allProductCategoryMap!=null){
            RPTProductCategory rptProductCategory = allProductCategoryMap.get(gradedOrderDataOfMQ.getProductCategoryId());
            if (rptProductCategory!=null){
                gradedOrderDataOfMQ.setProductCategoryName(rptProductCategory.getName());
            }else {
                gradedOrderDataOfMQ.setProductCategoryName("");
            }
        }else {
            gradedOrderDataOfMQ.setProductCategoryName("");
        }
//        gradedOrderDataOfMQ.setOrderItems(RPTOrderItemUtils.fromOrderItemsJson(gradedOrderDataOfMQ.getOrderItemJson()));
        gradedOrderDataOfMQ.setOrderItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(gradedOrderDataOfMQ.getOrderItemPb()));
        orderItemList.addAll(gradedOrderDataOfMQ.getOrderItems());
        //获取区域名
        RPTArea rptArea = areaMap.get(gradedOrderDataOfMQ.getCountyId());
        if (rptArea!=null) {
            RPTArea city = rptArea.getParent();
            if (city!=null){
                gradedOrderDataOfMQ.setCityId(city.getId());
                RPTArea cityArea = cityMap.get(city.getId());
                if (cityArea!=null){
                    RPTArea parent = cityArea.getParent();
                    if (parent!=null){
                        gradedOrderDataOfMQ.setProvinceId(parent.getId());
                    }
                }
            }
        }
        RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
        return gradedOrderDataOfMQ;
    }


    /**
     *保存消息队列的工单到中间表
     */
    public void saveGradeOrderOfMQToRptDB(MQRPTOrderProcessMessage.RPTOrderProcessMessage msg){
        if(msg.getOrderId()<=0){
            throw new RuntimeException("客评工单订单id不能为空");
        }
        if(TextUtils.isEmpty(msg.getQuarter())){
            throw new RuntimeException("客评工单订单分片不能为空");
        }
        RPTGradedOrderEntity gradedOrderOfMQ = getGradedOrderOfMQ(msg.getQuarter(), msg.getOrderId());
        if (gradedOrderOfMQ!=null) {
            int systemId = RptCommonUtils.getSystemId();
            int dayIndex = Integer.parseInt(DateUtils.getDay(gradedOrderOfMQ.getCloseDateD()));
            String quarter = DateUtils.getQuarter(gradedOrderOfMQ.getCloseDateD());
            Integer yearmonth = StringUtils.toInteger(DateUtils.getYearMonth(gradedOrderOfMQ.getCloseDateD()));
            insertGradedOrderRpt(gradedOrderOfMQ,quarter,dayIndex,systemId,yearmonth);
        }
    }

    /**
     *保存客评工单
     */
    public void saveGradedOrderToRptDB(Date queryDate){
        if (queryDate!=null){
            List<RPTGradedOrderEntity> gradedOrderData = getGradedOrderData(queryDate);
            if (!gradedOrderData.isEmpty()){
                String quarter = DateUtils.getQuarter(queryDate);
                int systemId = RptCommonUtils.getSystemId();
                int dayIndex = Integer.parseInt(DateUtils.getDay(queryDate));
                Integer yearmonth = StringUtils.toInteger(DateUtils.getYearMonth(queryDate));
                for (RPTGradedOrderEntity entity:gradedOrderData) {
                    insertGradedOrderRpt(entity,quarter,dayIndex,systemId,yearmonth);
                }
            }
        }
    }

    /**
     * 保存遗漏工单
     */
    public void saveMissGradedOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTGradedOrderEntity> list = getGradedOrderData(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                int dayIndex = Integer.parseInt(DateUtils.getDay(date));
                Date beginDate = DateUtils.getDateStart(date);
                Date endDate = DateUtils.getDateEnd(date);
                List<LongTwoTuple> tuples = gradedOrderRptMapper.getGradedOrderIds(quarter, systemId, beginDate.getTime(), endDate.getTime());
                Map<Long, Long> gradedOrderMap = new HashMap<>();
                if (tuples != null && !tuples.isEmpty()) {
                   gradedOrderMap = tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
                }
                Integer yearmonth = StringUtils.toInteger(DateUtils.getYearMonth(date));
                Long primaryKeyId;
                for (RPTGradedOrderEntity item : list) {
                    primaryKeyId = gradedOrderMap.get(item.getOrderId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        insertGradedOrderRpt(item, quarter, dayIndex,systemId,yearmonth);
                    }
                }
            }
        }
    }

    /**
     * 删除重复的工单
     */
    public void deleteHavingGradedOrder(Date date){
        if (date!=null){
            String quarter = QuarterUtils.getSeasonQuarter(date);
            int systemId = RptCommonUtils.getSystemId();
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            List<LongTwoTuple> havingGradedOrder = gradedOrderRptMapper.getHavingGradedOrder(quarter, systemId, beginDate.getTime(), endDate.getTime());

            for (LongTwoTuple twoTuple: havingGradedOrder) {
                Long bElement = twoTuple.getBElement();
                if (bElement>1){
                   gradedOrderRptMapper.deleteHavingGradedOrder(twoTuple.getAElement(),bElement.intValue()-1);
               }

            }
        }
    }


    /**
     *写入客评工单到中间表
     */
    public void insertGradedOrderRpt(RPTGradedOrderEntity entity,String quarter,int dayIndex,int systemId,int yearmonth){
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            entity.setDayIndex(dayIndex);
            entity.setYearmonth(yearmonth);
            StringBuilder stringBuilder = new StringBuilder();
            int index = 0;
            for (RPTOrderItem item : entity.getOrderItems()) {
                if (index != 0) {
                    stringBuilder.append(",");
                }
                index++;
                if (item.getProduct() != null) {
                    stringBuilder.append(item.getProduct().getName());
                }
            }
            entity.setProductNames(stringBuilder.toString());
            gradedOrderRptMapper.insertGradedOrder(entity);
        }catch (Exception e){
            log.error("GradedOrderRptService.insertGradedOrderRpt:{}", Exceptions.getStackTraceAsString(e));
        }
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
                            saveGradedOrderToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissGradedOrdersToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteGradedOrdersFromRptDB(beginDate);
                            saveGradedOrderToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteGradedOrdersFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("GradedOrderRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 删除中间表中指定日期的已客评工单
     */
    public void deleteGradedOrdersFromRptDB(Date date) {
        if (date != null) {
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            gradedOrderRptMapper.deleteGradedOrders(quarter, systemId, beginDate.getTime(), endDate.getTime());
        }
    }

    /**
     * 获取工单费用报表数据
     */
    public Page<RPTGradedOrderEntity> getOrderServicePointFeeOfGradedOrder(RPTGradedOrderSearch searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition.getPageNo() != null && searchCondition.getPageSize() != null) {
            PageHelper.startPage(searchCondition.getPageNo(), searchCondition.getPageSize());
        }
        Page<RPTGradedOrderEntity> orderServicePointFeeGradedOrderData = gradedOrderRptMapper.getOrderServicePointFeeGradedOrderData(searchCondition.getBeginDate(),
                searchCondition.getEndDate(),systemId, searchCondition.getServicePointId(), searchCondition.getQuarters(), searchCondition.getProductCategoryIds());

        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> countyMap = areaCacheService.getAllCountyMap();
        List<Long> servicePointIds = orderServicePointFeeGradedOrderData.stream().map(RPTGradedOrderEntity::getServicePointId).collect(Collectors.toList());
        List<Long> customerIds = orderServicePointFeeGradedOrderData.stream().map(RPTGradedOrderEntity::getCustomerId).collect(Collectors.toList());
        List<Long> distinctServicePointIds = servicePointIds.stream().distinct().collect(Collectors.toList());
        Set<Long> distinctCustomerIds = customerIds.stream().distinct().collect(Collectors.toSet());
        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(distinctCustomerIds), Arrays.asList(fieldsArray));
        Map<Long, RPTServicePoint> servicePointMap = msServicePointService.getServicePointMap(distinctServicePointIds);

        for (RPTGradedOrderEntity entity:orderServicePointFeeGradedOrderData) {
            if (TextUtils.isBlank(entity.getCustomerName())){
                if (customerMap!=null) {
                    RPTCustomer rptCustomer = customerMap.get(entity.getCustomerId());
                    if (rptCustomer != null) {
                        entity.setCustomerName(rptCustomer.getName());
                    }else {
                        entity.setCustomerName("");
                    }
                }
            }
            if (TextUtils.isBlank(entity.getServicePointNo()) || TextUtils.isBlank(entity.getServicePointName())){
                if (servicePointMap!=null) {
                    RPTServicePoint rptServicePoint = servicePointMap.get(entity.getServicePointId());
                    if (rptServicePoint != null) {
                        entity.setServicePointName(rptServicePoint.getName());
                        entity.setServicePointNo(rptServicePoint.getServicePointNo());
                    }
                }
            }
            if (entity.getProvinceId()!=null && provinceMap.get(entity.getProvinceId())!=null){
                entity.setProvinceName(provinceMap.get(entity.getProvinceId()).getName());
            }
            if (entity.getCityId()!=null && cityMap.get(entity.getCityId())!=null){
                entity.setCityName(cityMap.get(entity.getCityId()).getName());
            }
            if (entity.getCountyId()!=null && countyMap.get(entity.getCountyId())!=null){
                entity.setCountyName(countyMap.get(entity.getCountyId()).getName());
            }
            if (entity.getFirstPlanDate()!=null){
                entity.setFirstPlanDateD(DateUtils.longToDate(entity.getFirstPlanDate()));
            }
            if (entity.getPlanDate()!=null){
                entity.setPlanDateD(DateUtils.longToDate(entity.getPlanDate()));
            }
            if (entity.getCloseDate()!=null){
                entity.setCloseDateD(DateUtils.longToDate(entity.getCloseDate()));
            }
            if (entity.getCreateDate()!=null){
                entity.setCreateDateD(DateUtils.longToDate(entity.getCreateDate()));
            }
            if (entity.getArrivalDt()!=null){
                entity.setArrivalDate(DateUtils.longToDate(entity.getArrivalDt()));
            }
            if (entity.getArrivalDt()!=null){
                entity.setArrivalDate(DateUtils.longToDate(entity.getArrivalDt()));
            }
            if (entity.getAppCompleteDt() != null && entity.getAppCompleteDt() > 0){
                entity.setAppCompleteDate(DateUtils.longToDate(entity.getAppCompleteDt()));
            }
        }

        return orderServicePointFeeGradedOrderData;
    }


    /**
     * 获取客服每日完工数据
     * @param searchCondition
     * @return
     */
    public List<RPTKefuCompletedDailyEntity> getKefuCompletedDaily(RPTGradedOrderSearch searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        Set<Long> keFuIds = new HashSet<>();
        Map<Long, RPTUser> kAKeFuMap =  MSUserUtils.getMapByUserType(2);
        if(searchCondition.getSubFlag() != null && searchCondition.getSubFlag() != -1) {
            keFuUtils.getKeFu(kAKeFuMap,searchCondition.getSubFlag(),keFuIds);
        }

        List<Long> keFuIdsList =  Lists.newArrayList(keFuIds);
        Map<Long, List<RPTKefuCompletedDailyEntity>> kefusMap = new HashMap<>();
        List<RPTKefuCompletedDailyEntity> kefuGradedOrderDataList = gradedOrderRptMapper.getKefuGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(),systemId, searchCondition.getKefuId(),searchCondition.getProductCategoryIds(), searchCondition.getQuarter());
        Map<Long, List<RPTKefuCompletedDailyEntity>> kefuMap = kefuGradedOrderDataList.stream().collect(Collectors.groupingBy(RPTKefuCompletedDailyEntity::getKefuId));

        if(searchCondition.getSubFlag() != null && searchCondition.getSubFlag() != -1){
            for(Long id : keFuIdsList){
               if(kefuMap.containsKey(id)){
                   kefusMap.put(id,kefuMap.get(id));
               }
            }
        }else{
            kefusMap = kefuMap;
        }
        RPTUser user;
        List<RPTKefuCompletedDailyEntity> list = Lists.newArrayList();
        for (List<RPTKefuCompletedDailyEntity> completedDailyEntityList : kefusMap.values()) {
            RPTKefuCompletedDailyEntity completedDailyEntity = new RPTKefuCompletedDailyEntity();
            if (kAKeFuMap != null && kAKeFuMap.size() > 0) {
                user = kAKeFuMap.get(completedDailyEntityList.get(0).getKefuId());
                if (user != null) {
                    completedDailyEntity.setKefuName(user.getName());
                }
            }
            Double total = 0.0;
            Class countyItemClass = completedDailyEntity.getClass();
            for (RPTKefuCompletedDailyEntity entity : completedDailyEntityList) {
                total += entity.getCountSum();
                String strSetDMethodName = "setD" + entity.getDayIndex();
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Double.class);
                    setDMethod.invoke(completedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
            completedDailyEntity.setTotal(total);
            list.add(completedDailyEntity);
        }
        Date lastMonth = DateUtils.addMonth(new Date(searchCondition.getBeginDate()), -1);
        String quarter = QuarterUtils.getSeasonQuarter(lastMonth);
        Date startDayOfMonth = DateUtils.getStartDayOfMonth(lastMonth);
        Date lastDayOfMonth = DateUtils.getLastDayOfMonth(lastMonth);
        int lastMonthDays = DateUtils.getDaysOfMonth(lastMonth);
        Long lastMonthOrderCount = gradedOrderRptMapper.getCompletedOrderSum(startDayOfMonth.getTime(),lastDayOfMonth.getTime(),quarter,systemId,searchCondition.getProductCategoryIds(),keFuIdsList);

        lastMonthOrderCount = (lastMonthOrderCount == null ? 0 : lastMonthOrderCount);
        double lastAvgCount = lastMonthOrderCount / lastMonthDays;

        RPTKefuCompletedDailyEntity sumKFOCD = new RPTKefuCompletedDailyEntity();
        sumKFOCD.setRowNumber(RPT_ROW_NUMBER_SUMROW);
        sumKFOCD.setKefuName("总计(单)");
        RPTKefuCompletedDailyEntity perKFOCD = new RPTKefuCompletedDailyEntity();
        perKFOCD.setRowNumber(RPT_ROW_NUMBER_PERROW);
        perKFOCD.setKefuName("每日完工环比(%)");
        RPTBaseDailyEntity.computeSumAndPerForCount(list, lastAvgCount, lastMonthOrderCount, sumKFOCD, perKFOCD);
        RPTKefuCompletedDailyEntity perSKFD = new RPTKefuCompletedDailyEntity();
        Date lastYearSomeMonth = DateUtils.addMonth(new Date(searchCondition.getBeginDate()), -12);
        Date goLiveDate = RptCommonUtils.getGoLiveDate();
        if(goLiveDate.getTime() < lastYearSomeMonth.getTime()) {
            Date dayOfMonth = DateUtils.getStartDayOfMonth(lastYearSomeMonth);
            Date lastDayMonth = DateUtils.getLastDayOfMonth(lastYearSomeMonth);
            int lastYearSomeMonthDays = DateUtils.getDaysOfMonth(lastMonth);
            String lastYearQuarter = DateUtils.getQuarter(dayOfMonth);
            Long lastYearSomeMonthCount = gradedOrderRptMapper.getCompletedOrderSum(dayOfMonth.getTime(), lastDayMonth.getTime(), lastYearQuarter,systemId, searchCondition.getProductCategoryIds(),keFuIdsList);
            lastYearSomeMonthCount = (lastYearSomeMonthCount == null ? 0 : lastYearSomeMonthCount);
            double lastYearSomeMonthAvgCount = lastYearSomeMonthCount / lastYearSomeMonthDays;
            computeSumAndPerForCount(list, lastYearSomeMonthAvgCount, lastYearSomeMonthCount, sumKFOCD, perSKFD);
        }
        perSKFD.setRowNumber(RPT_ROW_NUMBER_PERROW);
        perSKFD.setKefuName("每日完工同比(%)");
        list.add(sumKFOCD);
        list.add(perKFOCD);
        list.add(perSKFD);

        return list;
    }
    public static void computeSumAndPerForCount(List baseDailyReports, double lastMonthAvgCount, double lastMonthTotalCount,
                                                RPTBaseDailyEntity sumDailyReport, RPTBaseDailyEntity perDailyReport) {
        //计算每月订单数的对比值
        if (sumDailyReport != null && perDailyReport != null) {
            Class sumDailyReportClass = sumDailyReport.getClass();
            Class perDailyReportClass = perDailyReport.getClass();
            for (int i = 1; i < 32; i++) {
                String strGetMethodName = "getD" + i;
                String strSetMethodName = "setD" + i;
                try {
                    Method sumDailyReportGetMethod = sumDailyReportClass.getMethod(strGetMethodName);
                    Object sumDailyReportGetD = sumDailyReportGetMethod.invoke(sumDailyReport);

                    double sumD = StringUtils.toDouble(sumDailyReportGetD);
                    double perD = -100;
                    if (lastMonthAvgCount != 0) {
                        perD = (sumD - lastMonthAvgCount) / lastMonthAvgCount * 100;
                    }

                    Method perDailyReportSetMethod = perDailyReportClass.getMethod(strSetMethodName, Double.class);
                    perDailyReportSetMethod.invoke(perDailyReport, perD);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }

            double perTotal = -100;
            if (lastMonthTotalCount != 0) {
                perTotal = (sumDailyReport.getTotal() - lastMonthTotalCount) / lastMonthTotalCount * 100;
            }
            perDailyReport.setTotal(perTotal);
        }
    }
    /**
     * 获取省每日完工数据
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getProvinceCompletedOrderData(RPTGradedOrderSearch searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        List<RPTAreaCompletedDailyEntity> provinceGradedOrderData = gradedOrderRptMapper.getProvinceGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(),systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );
        Map<Long, List<RPTAreaCompletedDailyEntity>> provinceCompletedMap = provinceGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getProvinceId));
        Map<Long, RPTArea> areaMap = areaCacheService.getAllProvinceMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();
        for (List<RPTAreaCompletedDailyEntity> entityList : provinceCompletedMap.values()) {
            RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            RPTArea rptArea = areaMap.get(entityList.get(0).getProvinceId());
            if (rptArea!=null){
                rptAreaCompletedDailyEntity.setProvinceId(rptArea.getId());
                rptAreaCompletedDailyEntity.setProvinceName(rptArea.getName());
            }
            Double total = 0.0;
            Class countyItemClass = rptAreaCompletedDailyEntity.getClass();
            for (RPTAreaCompletedDailyEntity entity : entityList) {
                total += Double.valueOf(entity.getCountSum());
                String strSetDMethodName = "setD" + entity.getDayIndex();
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Double.class);
                    setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
            rptAreaCompletedDailyEntity.setTotal(total);
            list.add(rptAreaCompletedDailyEntity);

        }
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setProvinceName("总计(单)");
        sumUp.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        list.add(sumUp);
        return list;
    }
    /**
     * 获取市每日完工数据
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getCityCompletedOrderData(RPTGradedOrderSearch searchCondition) {
        int systemId = RptCommonUtils.getSystemId();
        List<RPTAreaCompletedDailyEntity> cityGradedOrderData = gradedOrderRptMapper.getCityGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(), systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );
        Map<Long, List<RPTAreaCompletedDailyEntity>> cityCompletedMap = cityGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCityId));
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCityMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();

        for (List<RPTAreaCompletedDailyEntity> entityList : cityCompletedMap.values()) {
            RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            RPTArea rptArea = areaMap.get(entityList.get(0).getCityId());
            RPTArea rptProvince = provinceMap.get(entityList.get(0).getProvinceId());

            if (rptArea != null) {
                rptAreaCompletedDailyEntity.setCityId(rptArea.getId());
                rptAreaCompletedDailyEntity.setCityName(rptArea.getName());
            }


            if (rptProvince != null) {
                rptAreaCompletedDailyEntity.setProvinceId(rptProvince.getId());
                rptAreaCompletedDailyEntity.setProvinceName(rptProvince.getName());
            }

            Double total = 0.0;
            Class countyItemClass = rptAreaCompletedDailyEntity.getClass();
            for (RPTAreaCompletedDailyEntity entity : entityList) {
                total += Double.valueOf(entity.getCountSum());
                String strSetDMethodName = "setD" + entity.getDayIndex();
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Double.class);
                    setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
            rptAreaCompletedDailyEntity.setTotal(total);
            list.add(rptAreaCompletedDailyEntity);

        }
        list = list.stream().sorted(Comparator.comparing(RPTAreaCompletedDailyEntity::getProvinceId)).collect(Collectors.toList());
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setProvinceName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        list.add(sumUp);
        return list;
    }

    /**
     * 获取区域每日完工数据
     *
     * @param searchCondition
     * @return
     */
    public List<RPTAreaCompletedDailyEntity> getAreaCompletedOrderData(RPTGradedOrderSearch searchCondition) {
        int systemId = RptCommonUtils.getSystemId();
        List<RPTAreaCompletedDailyEntity> areaGradedOrderData = gradedOrderRptMapper.getAreaGradedOrderData(searchCondition.getBeginDate(), searchCondition.getEndDate(), systemId, searchCondition.getAreaId(),
                searchCondition.getAreaType(), searchCondition.getCustomerId(), searchCondition.getQuarter(), searchCondition.getProductCategoryIds()
        );
        Map<Long, List<RPTAreaCompletedDailyEntity>> areaCompletedMap = areaGradedOrderData.stream().collect(Collectors.groupingBy(RPTAreaCompletedDailyEntity::getCountyId));
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        List<RPTAreaCompletedDailyEntity> list = Lists.newArrayList();
        for (List<RPTAreaCompletedDailyEntity> entityList : areaCompletedMap.values()) {
            RPTAreaCompletedDailyEntity rptAreaCompletedDailyEntity = new RPTAreaCompletedDailyEntity();
            RPTArea rptProvince = provinceMap.get(entityList.get(0).getProvinceId());
            RPTArea rptCity = cityMap.get(entityList.get(0).getCityId());
            RPTArea rptArea = areaMap.get(entityList.get(0).getCountyId());
            if (rptProvince != null) {
                rptAreaCompletedDailyEntity.setProvinceId(rptProvince.getId());
                rptAreaCompletedDailyEntity.setProvinceName(rptProvince.getName());
            }

            if (rptCity != null) {
                rptAreaCompletedDailyEntity.setCityId(rptCity.getId());
                rptAreaCompletedDailyEntity.setCityName(rptCity.getName());
            }

            if (rptArea != null) {
                rptAreaCompletedDailyEntity.setCountyId(rptArea.getId());
                rptAreaCompletedDailyEntity.setCountyName(rptArea.getName());
            }

            Double total = 0.0;
            Class countyItemClass = rptAreaCompletedDailyEntity.getClass();
            for (RPTAreaCompletedDailyEntity entity : entityList) {
                total += Double.valueOf(entity.getCountSum());
                String strSetDMethodName = "setD" + entity.getDayIndex();
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Double.class);
                    setDMethod.invoke(rptAreaCompletedDailyEntity, StringUtils.toDouble(entity.getCountSum()));
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
            rptAreaCompletedDailyEntity.setTotal(total);
            list.add(rptAreaCompletedDailyEntity);

        }
        list = list.stream().sorted(Comparator.comparing(RPTAreaCompletedDailyEntity::getProvinceId)
                .thenComparing(RPTAreaCompletedDailyEntity::getCityId)).collect(Collectors.toList());
        RPTAreaCompletedDailyEntity sumUp = new RPTAreaCompletedDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setCountyName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(list, 0, 0, sumUp, null);
        list.add(sumUp);
        return list;
    }

    /**
     * 获取开发均单费用报表数据
     * @param search
     * @return
     */
    public List<RPTDevelopAverageOrderFeeEntity> getDevelopAverageOrderFee(RPTGradedOrderSearch search){
        int systemId = RptCommonUtils.getSystemId();
        List<RPTDevelopAverageOrderFeeEntity> list = gradedOrderRptMapper.getDevelopAverageOrderFee(search.getBeginDate(), search.getEndDate(), systemId, search.getQuarter(), search.getProductCategoryIds());
        Map<Long, List<RPTDevelopAverageOrderFeeEntity>> createByMap = list.stream().collect(Collectors.groupingBy(RPTDevelopAverageOrderFeeEntity::getCrushCreateBy));
        Set<Long> createByIdSet = createByMap.keySet();
        List<Long> createByList = new ArrayList<>(createByIdSet);
        Map<Long, String> namesByUserIds = MSUserUtils.getNamesByUserIds(createByList);
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        List<RPTDevelopAverageOrderFeeEntity> emptyList = Lists.newArrayList();
        int totalQty = 0;
        int total96Hour = 0;
        double totalMakeCharge = 0.0;
        for (List<RPTDevelopAverageOrderFeeEntity> entityList : createByMap.values()) {
            RPTDevelopAverageOrderFeeEntity  feeEntity = new RPTDevelopAverageOrderFeeEntity();
            int qty = entityList.size();
            feeEntity.setCompletedQty(qty);
            int order96HourCompletedQty = 0;
            double makeCharge = 0.0;
            Set<String> productCategoryNameList = new HashSet<>();

            for (RPTDevelopAverageOrderFeeEntity entity:entityList) {
                if (entity.getOrder96HourCompletedFlag()==1){
                    order96HourCompletedQty += 1;
                }
                makeCharge = makeCharge + entity.getEngineerOtherCharge() + entity.getEngineerTravelCharge();
                productCategoryNameList.add(entity.getProductCategoryNames());
            }
            feeEntity.setOrder96HourCompletedQty(order96HourCompletedQty);
            feeEntity.setMakeCharge(makeCharge);
            feeEntity.setAverageCharge(makeCharge / qty);
            String productCateNames = String.join(",", productCategoryNameList);
            feeEntity.setProductCategoryNames(productCateNames);
            String format = numberFormat.format((float) order96HourCompletedQty / qty * 100);
            feeEntity.setOrder96HourCompletedRate(format);
            feeEntity.setCrushCreateBy(entityList.get(0).getCrushCreateBy());
            feeEntity.setCloseDateDt(DateUtils.longToDate(entityList.get(0).getCloseDate()));
            feeEntity.setCrushCreateByName(namesByUserIds.get(entityList.get(0).getCrushCreateBy()));
            emptyList.add(feeEntity);
            totalQty = totalQty + qty;
            total96Hour= total96Hour + order96HourCompletedQty;
            totalMakeCharge = totalMakeCharge + makeCharge;
        }
        if (emptyList.size()>0) {
            RPTDevelopAverageOrderFeeEntity sumEntity = new RPTDevelopAverageOrderFeeEntity();
            sumEntity.setCompletedQty(totalQty);
            sumEntity.setOrder96HourCompletedQty(total96Hour);
            sumEntity.setMakeCharge(totalMakeCharge);
            sumEntity.setAverageCharge(totalMakeCharge / totalQty);
            String format = numberFormat.format((float) total96Hour / totalQty * 100);
            sumEntity.setOrder96HourCompletedRate(format);
            emptyList.add(sumEntity);
        }
        return emptyList;
    }

    /**
     * 检查客服每日完工报表是否有数据存在
     */
    public boolean hasKefuCompletedOrderReportData(String searchConditionJson) {
        boolean result = false;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        Map<Long, RPTUser> kAKeFuMap;
        Set<Long> keFuIds = new HashSet<>();
        if (searchCondition != null) {
            if(searchCondition.getSubFlag() != null && searchCondition.getSubFlag() != -1) {
                result = true;
                kAKeFuMap = MSUserUtils.getMapByUserType(2);
                keFuUtils.getKeFu(kAKeFuMap,searchCondition.getSubFlag(),keFuIds);
            }
            List<Long> keFuIdsList =  Lists.newArrayList();
            List<Long> keFuSubTypeList =  Lists.newArrayList(keFuIds);
            if(searchCondition.getKefuId() != 0 && result){
                if(keFuSubTypeList.size() > 0){
                    for(Long id : keFuSubTypeList){
                        if(searchCondition.getKefuId().equals(id)){
                            keFuIdsList.add(id);
                        }
                    }
                    if(keFuIdsList.size() == 0){
                        return false;
                    }
                }else {
                    return false;
                }
            }else if(searchCondition.getKefuId() != 0){
                keFuIdsList.add(searchCondition.getKefuId());
            }else if(result){
                if(keFuSubTypeList.size() > 0){
                    keFuIdsList.addAll(keFuSubTypeList);
                }else {
                    return false;
                }
            }
            Integer rowCount = gradedOrderRptMapper.hasKefuCompletedOrderReportData(searchCondition.getBeginDate(),
                    searchCondition.getEndDate(),keFuIdsList,systemId,searchCondition.getProductCategoryIds()
                    ,searchCondition.getQuarter(),searchCondition.getQuarters());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 检查区域每日完工报表是否有数据存在
     */
    public boolean hasAreaCompletedOrderReportData(String searchConditionJson) {
        boolean result = true;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition != null) {
            Integer rowCount = gradedOrderRptMapper.hasAreaCompletedOrderReportData(searchCondition.getBeginDate(),
                    searchCondition.getEndDate(),searchCondition.getAreaType(),searchCondition.getAreaId(),systemId,searchCondition.getCustomerId()
                    ,searchCondition.getQuarter(),searchCondition.getProductCategoryIds(),searchCondition.getQuarters());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 检查工单费用报表是否有数据存在
     */
    public boolean hasOrderServicePointFeeReportData(String searchConditionJson) {
        boolean result = true;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition != null) {
            Integer rowCount = gradedOrderRptMapper.hasOrderServicePointFeeReportData(searchCondition.getBeginDate(),
                    searchCondition.getEndDate(),systemId,searchCondition.getServicePointId()
                    ,searchCondition.getQuarter(),searchCondition.getProductCategoryIds(),searchCondition.getQuarters());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 检查开发均单费用报表是否有数据存在
     */
    public boolean hasDevelopAverageFeeReportData(String searchConditionJson) {
        boolean result = true;
        RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
        int systemId = RptCommonUtils.getSystemId();
        if (searchCondition != null) {
            Integer rowCount = gradedOrderRptMapper.hasDevelopAverageFeeReportData(searchCondition.getBeginDate(),
                    searchCondition.getEndDate(),systemId,searchCondition.getQuarter()
                    ,searchCondition.getProductCategoryIds(),searchCondition.getQuarters());
            result = rowCount > 0;
        }
        return result;
    }


    /**
     *创建区县每日完工报表
     */
    public SXSSFWorkbook exportCountyCompletedRpt(String searchConditionJson, String reportTitle){
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 3));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, days + 2));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日下单(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 3, days + 3));
            ExportExcel.createCell(headFirstRow, days + 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");
            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
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
                    pColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, province, pColumnIndex, provinceClass);
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotal());
                    if (cityList != null && cityList.size() > 0) {
                        for (int j = 0; j < cityList.size() - 1; j++) {
                            RPTAreaCompletedDailyEntity city = cityList.get(j);
                            if (city.getProvinceId().intValue() == province.getProvinceId().intValue()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                int cColumnIndex = 0;
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, null == province.getProvinceName() ? "" : province.getProvinceName());
                                dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                                Class cityClass = city.getClass();
                                cColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, city, cColumnIndex, cityClass);
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotal());

                                //循环读取市下面的区
                                if (areaList != null && areaList.size() > 0) {
                                    for (int k = 0; k < areaList.size() - 1; k++) {
                                        RPTAreaCompletedDailyEntity area = areaList.get(k);

                                        if (area.getCityId().intValue() == city.getCityId().intValue()) {
                                            dataRow = xSheet.createRow(rowIndex++);
                                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                            int aColumnIndex = 0;

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, null == province.getProvinceName() ? "" : province.getProvinceName());

                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, null == city.getCityName() ? "" : city.getCityName());
                                            dataCell = ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                            dataCell.setCellValue(null == area.getCountyName() ? "" : area.getCountyName());
                                            Class areaClass = area.getClass();

                                            aColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, area, aColumnIndex, areaClass);
                                            ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, area.getTotal());
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

                columnIndex = writeDailyPlanOrders(days, xStyle, dataRow, sumAOPD, columnIndex, totalClass);

                Double totalCount = StringUtils.toDouble(sumAOPD.getTotal());
                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);
            }
        }

        catch (Exception e) {
            log.error("【GradedOrderRptService.exportCountyCompletedRpt】区每日完工表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;

    }
        return xBook;

    }

    private int writeDailyPlanOrders(int days, Map<String, CellStyle> xStyle, Row dataRow, RPTAreaCompletedDailyEntity entity, int pColumnIndex, Class provinceClass) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException
            {
                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                    String strGetMethodName = "getD" + dayIndex;

                    Method method = provinceClass.getMethod(strGetMethodName);
                    Object objGetD = method.invoke(entity);
                    Double d = com.kkl.kklplus.utils.StringUtils.toDouble(objGetD);
                    String strD = String.format("%.0f", d);

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                }
                return pColumnIndex;
            }

    /**
     * 创建市每日完工报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportCityCompletedRpt(String searchConditionJson, String reportTitle){
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 2));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, days+1));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 2, days + 2));
            ExportExcel.createCell(headFirstRow, days + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex+1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
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
                    pColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, province, pColumnIndex, provinceClass);
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, province.getTotal());
                    if (cityList != null && cityList.size() > 0) {
                        for (int j = 0; j < cityList.size() - 1; j++) {
                            RPTAreaCompletedDailyEntity city = cityList.get(j);
                            if (city.getProvinceId().intValue() == province.getProvinceId().intValue()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                int cColumnIndex = 0;
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, null == province.getProvinceName() ? "" : province.getProvinceName());
                                dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                                dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());
                                Class cityClass = city.getClass();
                                cColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, city, cColumnIndex, cityClass);
                                ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getTotal());
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

                columnIndex = writeDailyPlanOrders(days, xStyle, dataRow, sumAOPD, columnIndex, totalClass);

                Double totalCount = StringUtils.toDouble(sumAOPD.getTotal());
                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);
            }
        }
        catch (Exception e) {
            log.error("【GradedOrderRptService.exportCityCompletedRpt】市每日完工写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

    /**
     * 创建省每日完工报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportProvinceCompletedRpt(String searchConditionJson, String reportTitle){
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 1));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, days));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 1, days + 1));
            ExportExcel.createCell(headFirstRow, days + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            int index = 0;
            if (list != null && list.size() > 0) {
                for (RPTAreaCompletedDailyEntity entity : list) {
                    index ++;
                    if (list.size()>index){
                        dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int pColumnIndex = 0;
                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                        Class provinceClass = entity.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetMethodName = "getD" + dayIndex;

                            Method method = provinceClass.getMethod(strGetMethodName);
                            Object objGetD = method.invoke(entity);
                            Double d = StringUtils.toDouble(objGetD);
                            String strD = String.format("%.0f", d);

                            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                        }
                        ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotal());

                    }else {
                        //读取总计
                        dataRow = xSheet.createRow(rowIndex);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 1;

                        dataCell = ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                        dataCell.setCellValue(null == entity.getProvinceName() ? "" : entity.getProvinceName());

                        Class totalClass = entity.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetMethodName = "getD" + dayIndex;

                            Method method = totalClass.getMethod(strGetMethodName);
                            Object objGetD = method.invoke(entity);
                            Double d = StringUtils.toDouble(objGetD);
                            String strD = String.format("%.0f", d);
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                        }
                        Double totalCount = StringUtils.toDouble(entity.getTotal());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);
                    }

                }
            }
        }
        catch (Exception e) {
            log.error("【GradedOrderRptService.exportProvinceCompletedRpt】省每日完工写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 创建客服每日完工报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportKefuCompletedRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            List<RPTKefuCompletedDailyEntity> list = getKefuCompletedDaily(searchCondition);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getBeginDate()));
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 1));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, days));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日完工单(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 1, days + 1));
            ExportExcel.createCell(headFirstRow, days + 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKefuCompletedDailyEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == rowData.getKefuName() ? "" : rowData.getKefuName().toString());

                    Class rowDataClass = rowData.getClass();
                    for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                        String strGetMethodName = "getD" + dayIndex;

                        Method method = rowDataClass.getMethod(strGetMethodName);
                        Object objGetD = method.invoke(rowData);
                        Double d = StringUtils.toDouble(objGetD);
                        String strD = null;
                        if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                            strD = String.format("%.2f%s", d, '%');
                        } else {
                            strD = String.format("%.0f", d);
                        }
                        dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
                    }

                    Double totalCount = StringUtils.toDouble(rowData.getTotal());
                    String strTotalCount = null;
                    if (rowData.getRowNumber() == RPTBaseDailyEntity.RPT_ROW_NUMBER_PERROW) {
                        strTotalCount = String.format("%.2f%s", totalCount, '%');
                        ;
                    } else {
                        strTotalCount = String.format("%.0f", totalCount);
                    }
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
                }
            }

        }catch (Exception e){
            log.error("【GradedOrderRptService.exportKefuCompletedRpt】客服每日完工写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 创建工单费用报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportOrderServicePointFeeRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            Page<RPTGradedOrderEntity> list = getOrderServicePointFeeOfGradedOrder(searchCondition);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //===============绘制标题======================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 33));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单号");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户姓名");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务品类");
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单描述");
            ExportExcel.createCell(headFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            ExportExcel.createCell(headFirstRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "首次派单时间");
            ExportExcel.createCell(headFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单时间");
            ExportExcel.createCell(headFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货时间");
            ExportExcel.createCell(headFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "APP完成时间");
            ExportExcel.createCell(headFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评时间");
            ExportExcel.createCell(headFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headFirstRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
            ExportExcel.createCell(headFirstRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "材料费");
            ExportExcel.createCell(headFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
            ExportExcel.createCell(headFirstRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
            ExportExcel.createCell(headFirstRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
            ExportExcel.createCell(headFirstRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "互助基金");
            ExportExcel.createCell(headFirstRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快可立时效费");
            ExportExcel.createCell(headFirstRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户时效费");
            ExportExcel.createCell(headFirstRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headFirstRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");
            ExportExcel.createCell(headFirstRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台费");
            ExportExcel.createCell(headFirstRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保金额");
            ExportExcel.createCell(headFirstRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headFirstRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "总金额");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            // 写入数据

            if (list != null && list.size() > 0) {

                double serviceCharge = 0.0;
                double materialCharge = 0.0;
                double travelCharge = 0.0;
                double expressCharge = 0.0;
                double otherCharge = 0.0;
                double insuranceCharge = 0.0;
                double timeLinensCharge = 0.0;
                double customerTimeLinensCharge = 0.0;
                double urgentCharge = 0.0;
                double taxFee = 0.0;
                double infoFee = 0.0;
                double engineerDeposit = 0.0;
                double praiseFee = 0.0;
                double total = 0.0;
                int rowsCount = list.size();
                int rowNumber = 0;
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTGradedOrderEntity entity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    rowNumber++;
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderNo());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCustomerName());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCityName());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCountyName());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServiceAddress());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getUserName());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServicePhone());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProductCategoryName());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProductNames());

                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getDescription());
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(new Date(entity.getCreateDate()), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(new Date(entity.getFirstPlanDate()), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(new Date(entity.getPlanDate()), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getArrivalDate() == null ? "" : DateUtils.formatDate(entity.getArrivalDate(),"yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAppCompleteDate() == null ? "" : DateUtils.formatDate(entity.getAppCompleteDate(),"yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(new Date(entity.getCloseDate()), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServicePointNo());
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServicePointName());
                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServiceCharge());
                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getMaterialCharge());
                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTravelCharge());
                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getExpressCharge());
                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOtherCharge());
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getInsuranceCharge());
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTimeLinessCharge());
                    ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCustomerTimeLinessCharge());
                    ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getUrgentCharge());
                    ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTaxFee());
                    ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getInfoFee());
                    ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getEngineerDeposit());
                    ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getPraiseFee());
                    ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getTotalCharge());
                    serviceCharge = serviceCharge + entity.getServiceCharge();
                    materialCharge = materialCharge + entity.getMaterialCharge();
                    travelCharge = travelCharge + entity.getTravelCharge();
                    expressCharge = expressCharge + entity.getExpressCharge();
                    otherCharge = otherCharge + entity.getOtherCharge();
                    insuranceCharge = insuranceCharge + entity.getInsuranceCharge();
                    timeLinensCharge = timeLinensCharge + entity.getTimeLinessCharge();
                    customerTimeLinensCharge = customerTimeLinensCharge + entity.getCustomerTimeLinessCharge();
                    urgentCharge = urgentCharge + entity.getUrgentCharge();
                    taxFee = taxFee + entity.getTaxFee();
                    infoFee = infoFee + entity.getInfoFee();
                    engineerDeposit = engineerDeposit + entity.getEngineerDeposit();
                    praiseFee = praiseFee + entity.getPraiseFee();
                    total = total + entity.getTotalCharge();
                }
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 19));

                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge);
                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, materialCharge);
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, travelCharge);
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expressCharge);
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, otherCharge);
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, insuranceCharge);
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, timeLinensCharge);
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimeLinensCharge);
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, urgentCharge);
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, taxFee);
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, infoFee);
                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, total);
            }

        } catch (Exception e) {
            log.error("【GradedOrderRptService.exportOrderServicePointFeeRpt】工单费用写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 开发均单费用报表导出
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportDevelopAverageFeeRpt(String searchConditionJson, String reportTitle){
        SXSSFWorkbook xBook = null;
        try {
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            List<RPTDevelopAverageOrderFeeEntity> list = getDevelopAverageOrderFee(searchCondition);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //===============绘制标题======================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "部门");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "人员");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "当日处理完成");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "使用费用");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "均单费用");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "96小时完成");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "96小时完成率");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            // 写入数据

            if (list != null && list.size() > 0) {

                int totalCompletedQty = 0;
                double totalMakeCharge = 0.0;
                int totalOrder96CompletedQty = 0;
                int rowsCount = list.size();
                int rowNumber = 0;
                int index = 0;
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTDevelopAverageOrderFeeEntity entity = list.get(dataRowIndex);
                    index++;
                    if (rowsCount > index) {
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        rowNumber++;
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(entity.getCloseDateDt(), "yyyy-MM-dd"));
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProductCategoryNames());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCrushCreateByName());
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCompletedQty());
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getMakeCharge());
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f", entity.getAverageCharge()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrder96HourCompletedQty());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrder96HourCompletedRate()+"%");
                        totalCompletedQty = totalCompletedQty + entity.getCompletedQty();
                        totalOrder96CompletedQty = totalOrder96CompletedQty + entity.getOrder96HourCompletedQty();
                        totalMakeCharge = totalMakeCharge + entity.getMakeCharge();
                    } else {
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 3));

                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCompletedQty());
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getMakeCharge());
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f", entity.getAverageCharge()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrder96HourCompletedQty());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrder96HourCompletedRate()+"%");
                    }
                }
            }
        }catch (Exception e) {
            log.error("【GradedOrderRptService.exportDevelopAverageFeeRpt】开发均单费用报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}
