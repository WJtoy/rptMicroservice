package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCreatedOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CreatedOrderMapper;
import com.kkl.kklplus.provider.rpt.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
public class CreatedOrderService extends RptBaseService{

    @Autowired
    private CreatedOrderMapper createdOrderMapper;



    /**
     * 获取订单Id
     */
    public List<Long> findOrderIdList(Long startDt,Long endDt,String quarter){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        int systemId = RptCommonUtils.getSystemId();
        List<Long> allOrderIdList = Lists.newArrayList();
        List<Long> orderIdList = Lists.newArrayList();
        orderIdList = createdOrderMapper.findOrderIdList(startDt,endDt,pageNum,pageSize,quarter,systemId);
        if(orderIdList !=null && orderIdList.size()>0){
            allOrderIdList.addAll(orderIdList);
        }
        listSize = orderIdList.size();
        pageNum = pageNum + listSize;
        while (listSize == pageSize){
            orderIdList = createdOrderMapper.findOrderIdList(startDt,endDt,pageNum,pageSize,quarter,systemId);
            pageNum = pageNum + orderIdList.size();
            listSize = orderIdList.size();
            if(orderIdList.size()>0){
                allOrderIdList.addAll(orderIdList);
            }
        }
        return allOrderIdList;
    }

    /**
     * 获取某天web数据库的所有订单Id
     */
    public List<Long> findOrderIdListFromWebDB(Date startDate,Date endDate,String quarter){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        List<Long> allOrderIdList = Lists.newArrayList();
        List<Long> orderIdList = Lists.newArrayList();
        orderIdList = createdOrderMapper.findOrderIdListFromWebDB(startDate,endDate,pageNum,pageSize,quarter);
        if(orderIdList !=null && orderIdList.size()>0){
            allOrderIdList.addAll(orderIdList);
        }
        listSize = orderIdList.size();
        pageNum = pageNum + listSize;
        while (listSize == pageSize){
            orderIdList = createdOrderMapper.findOrderIdListFromWebDB(startDate,endDate,pageNum,pageSize,quarter);
            pageNum = pageNum + orderIdList.size();
            listSize = orderIdList.size();
            if(orderIdList.size()>0){
                allOrderIdList.addAll(orderIdList);
            }
        }
        return allOrderIdList;
    }

    /**
     * 根据时间从web数据库分页获取订单信息
     */
    public List<RPTCreatedOrderEntity> findOrderListByDate(Date startDate,Date endDate,String quarter){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        List<RPTCreatedOrderEntity> allCreatedOrder = Lists.newArrayList();
        List<RPTCreatedOrderEntity> createdOrderList = Lists.newArrayList();
        createdOrderList = createdOrderMapper.findOrderListByDate(startDate,endDate,pageNum,pageSize,quarter);
        if(createdOrderList !=null && createdOrderList.size()>0){
            allCreatedOrder.addAll(createdOrderList);
        }
        listSize = createdOrderList.size();
        pageNum = pageNum + createdOrderList.size();
        while(listSize == pageSize){
            createdOrderList = createdOrderMapper.findOrderListByDate(startDate,endDate,pageNum,pageSize,quarter);
            pageNum = pageNum + createdOrderList.size();
            listSize = createdOrderList.size();
            if(createdOrderList!=null && createdOrderList.size()>0){
                allCreatedOrder.addAll(createdOrderList);
            }
        }
        return allCreatedOrder;
    }



    /**
     * 添加数据
     */
    public void insert(RPTCreatedOrderEntity createdOrderEntity){
        try {
            createdOrderMapper.insert(createdOrderEntity);
        }catch (Exception e){
           throw new RuntimeException("【CreatedOrderService.insert】OrderId: {}, errorMsg: {}" +  "订单号:" + createdOrderEntity.getOrderId() +"错误原因:" + Exceptions.getStackTraceAsString(e));
        }

    }

    /**
     * 添加数据
     */
    @Transactional()
    public void processRptCreatedOrderMessage(MQRPTOrderProcessMessage.RPTOrderProcessMessage msg){
        int dayIndex;
        String strDayIndex;
        Integer yearmonth =0;
        String strYearmonth = "";
        RPTCreatedOrderEntity entity = new RPTCreatedOrderEntity();
        if(msg.getOrderId()<=0){
            throw new RuntimeException("保存每日下单订单号不能为空");
        }
        if(msg.getProvinceId()<=0){
            throw new RuntimeException("保存每日下单省份不能为空");
        }
        if(msg.getCityId()<=0){
            throw new RuntimeException("保存每日下单城市不能为空");
        }
        if(msg.getCountyId()<=0){
            throw new RuntimeException("保存每日下单区县不能为空");
        }
        if(msg.getCustomerId()<=0){
            throw new RuntimeException("保存每日下单客户不能为空");
        }
        if(msg.getKeFuId()<=0){
            throw new RuntimeException("保存每日下单客服不能为空");
        }
        if(msg.getProductCategoryId()<=0){
            throw new RuntimeException("保存每日下单品类不能为空");
        }
        if(msg.getOrderCreateDate()<=0){
            throw new RuntimeException("保存每日下单时间不能为空");
        }
        entity.setSystemId(RptCommonUtils.getSystemId());
        entity.setOrderId(msg.getOrderId());
        entity.setProvinceId(msg.getProvinceId());
        entity.setCityId(msg.getCityId());
        entity.setCountyId(msg.getCountyId());
        entity.setCustomerId(msg.getCustomerId());
        entity.setKeFuId(msg.getKeFuId());
        entity.setDataSource(msg.getDataSource());
        entity.setProductCategoryId(msg.getProductCategoryId());
        entity.setOrderCreateDt(msg.getOrderCreateDate());
        entity.setQuarter(QuarterUtils.getSeasonQuarter(entity.getOrderCreateDt()));
        strDayIndex = DateUtils.getDay(new Date(entity.getOrderCreateDt()));
        dayIndex = Integer.parseInt(strDayIndex);
        entity.setDayIndex(dayIndex);
        strYearmonth = DateUtils.getYearMonth(new Date(entity.getOrderCreateDt()));
        yearmonth = Integer.parseInt(strYearmonth);
        entity.setYearmonth(yearmonth);
        entity.setOrderServiceType(msg.getOrderServiceType());
        insert(entity);
    }

    /**
     * 根据时间和分片删除数据
     */
    public void deleteByDate(Long startDt,Long endDt,String quarter){
        try {
            Integer systemId = RptCommonUtils.getSystemId();
            createdOrderMapper.deleteByDate(startDt,endDt,quarter,systemId);
        }catch (Exception e){
            throw new RuntimeException("【CreatedOrderService.deleteByDate】OrderId: {}, errorMsg: {},删除每日下单失败,失败原因:" + e.getMessage());
        }

    }


    /**
     * 检查数据
     * 1.去掉重复的数据
     */
    public void deleteRepeatOrder(Date date){
         Date startDate = DateUtils.getDateStart(date);
         Date endDate = DateUtils.getDateEnd(date);
         Long startDt = startDate.getTime();
         Long endDt = endDate.getTime();
         String quarter = QuarterUtils.getSeasonQuarter(date);
         List<Long> orderIdList = findOrderIdList(startDt,endDt,quarter);
         if(orderIdList !=null && orderIdList.size()>0){
             Set<Long> set = new HashSet<>(orderIdList);
             //获得list与set的差集
             Collection rs = CollectionUtils.disjunction(orderIdList,set);
             //将collection转换为list
             List<Long> deleteList = new ArrayList<>(rs);
             if(deleteList!=null && deleteList.size()>0){
                 Integer systemId = RptCommonUtils.getSystemId();
                 for(Long orderId:deleteList){
                     TwoTuple<Long,String>  twoTuple= createdOrderMapper.getIdByOrderId(orderId,systemId);
                     if(twoTuple!=null && twoTuple.getAElement() !=null && twoTuple.getAElement()>0 && StringUtils.isNotBlank(twoTuple.getBElement())){
                         createdOrderMapper.delete(twoTuple.getAElement(),twoTuple.getBElement());
                     }
                 }
             }
         }
    }

    /**
     * 检查数据
     * 补充缺少的数据
     */
    public void replenishOrder(Date date){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        Long startDt = startDate.getTime();
        Long endDt = endDate.getTime();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<Long> orderIdList = findOrderIdList(startDt,endDt,quarter);
        List<RPTCreatedOrderEntity> createdOrderEntities = Lists.newArrayList();
        List<Long> orderIdListFromWebDB = Lists.newArrayList();
        Map<Long,RPTCreatedOrderEntity> createdOrderEntityMap = Maps.newHashMap();
        Integer systemId = RptCommonUtils.getSystemId();
        do{
            createdOrderEntities = createdOrderMapper.findOrderListByDate(startDate,endDate,pageNum,pageSize,quarter);
            if(createdOrderEntities !=null && createdOrderEntities.size()>0){
                orderIdListFromWebDB = createdOrderEntities.stream().map(RPTCreatedOrderEntity::getOrderId).collect(Collectors.toList());
                orderIdListFromWebDB.removeAll(orderIdList);
                if(orderIdListFromWebDB!=null && orderIdListFromWebDB.size()>0){
                    createdOrderEntityMap = createdOrderEntities.stream().collect(Collectors.toMap(RPTCreatedOrderEntity::getOrderId, a -> a,(k1,k2)->k1));
                    for(Long orderId:orderIdListFromWebDB){
                        RPTCreatedOrderEntity createdOrderEntity = createdOrderEntityMap.get(orderId);
                        if(createdOrderEntity !=null){
                            if(createdOrderEntity.getOrderCreateDate() !=null){
                                String strDayIndex = DateUtils.getDay(createdOrderEntity.getOrderCreateDate());
                                Integer dayIndex = StringUtils.toInteger(strDayIndex);
                                createdOrderEntity.setDayIndex(dayIndex);
                                String strYearMonth = DateUtils.getYearMonth(createdOrderEntity.getOrderCreateDate());
                                Integer yearMonth = Integer.parseInt(strYearMonth);
                                createdOrderEntity.setYearmonth(yearMonth);
                                createdOrderEntity.setOrderCreateDt(createdOrderEntity.getOrderCreateDate().getTime());
                            }
                            createdOrderEntity.setQuarter(quarter);
                            createdOrderEntity.setSystemId(systemId);
                            insert(createdOrderEntity);
                        }
                    }
                }
            }
            pageNum = pageNum + createdOrderEntities.size();
            listSize = createdOrderEntities.size();
        }while (listSize == pageSize);
    }

    /**
     * 重建中间表添加
     */
     public void saveRebuildData(Date date){
         int pageNum = 0;
         int pageSize = 5000;
         int listSize = 0;
         Date startDate = DateUtils.getDateStart(date);
         Date endDate = DateUtils.getDateEnd(date);
         String quarter = QuarterUtils.getSeasonQuarter(date);
         List<RPTCreatedOrderEntity>  createdOrderEntities = Lists.newArrayList();
         Integer systemId = RptCommonUtils.getSystemId();
         do{
             createdOrderEntities = createdOrderMapper.findOrderListByDate(startDate,endDate,pageNum,pageSize,quarter);
             if(createdOrderEntities !=null && createdOrderEntities.size()>0){
                 for(RPTCreatedOrderEntity item:createdOrderEntities){
                     if(item.getOrderCreateDate()!=null){
                         item.setOrderCreateDt(item.getOrderCreateDate().getTime());
                         item.setDayIndex(StringUtils.toInteger(DateUtils.getDay(item.getOrderCreateDate())));
                         item.setYearmonth(StringUtils.toInteger(DateUtils.getYearMonth(item.getOrderCreateDate())));
                     }
                     item.setQuarter(quarter);
                     item.setSystemId(systemId);
                     insert(item);
                 }
             }
             pageNum = pageNum + createdOrderEntities.size();
             listSize = createdOrderEntities.size();
         }while (listSize == pageSize);
     }
    /**
     * 重建中间表添加
     */
    public void deleteRebuildDate(Date date){
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        deleteByDate(startDate.getTime(),endDate.getTime(),quarter);
    }



    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt){
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveRebuildData(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            replenishOrder(beginDate);
                            break;
                        case UPDATE:
                            deleteRebuildDate(beginDate);
                            saveRebuildData(beginDate);
                            break;
                        case DELETE:
                            deleteRebuildDate(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CreatedOrderService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

}
