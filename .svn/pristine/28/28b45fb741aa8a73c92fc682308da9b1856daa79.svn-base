package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderDailyEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.CancelledOrderDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
public class CancelledOrderDailyRptService {

    @Autowired
    private CancelledOrderDailyRptMapper cancelledOrderDailyRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;


    /**
     * 保存接收消息队列退单取消单信息
     */
    public void save(MQRPTOrderProcessMessage.RPTOrderProcessMessage msg){
        int dayIndex;
        String strDayIndex;
        RPTCancelledOrderDailyEntity cancelledOrderDaily = new RPTCancelledOrderDailyEntity();
        if(msg.getOrderId()<=0){
            throw new RuntimeException("退单/取消单订单id不能为空");
        }
        if(msg.getCustomerId()<=0){
            throw new RuntimeException("退单/取消单客户id不能为空");
        }
        if(msg.getKeFuId()<=0){
            throw new RuntimeException("退单/取消单客服id不能为空");
        }
        if(msg.getOrderStatus()!=90 && msg.getOrderStatus()!=100){
            throw new RuntimeException("订单状态不是退单或者取消单");
        }
        if(msg.getOrderCreateDate()<=0){
            throw new RuntimeException("退单/取消单创建时间不能为空");
        }
        if(msg.getOrderCloseDate()<=0){
            throw new RuntimeException("退单/取消单关闭时间不能为空");
        }
        cancelledOrderDaily.setOrderId(msg.getOrderId());
        cancelledOrderDaily.setCustomerId(msg.getCustomerId());
        cancelledOrderDaily.setKeFuId(msg.getKeFuId());
        cancelledOrderDaily.setProductCategoryId(msg.getProductCategoryId());
        cancelledOrderDaily.setStatus(msg.getOrderStatus());
        cancelledOrderDaily.setOrderCreateDt(msg.getOrderCreateDate());
        cancelledOrderDaily.setOrderCloseDt(msg.getOrderCloseDate());
        cancelledOrderDaily.setDataSource(msg.getDataSource());
        cancelledOrderDaily.setQuarter(QuarterUtils.getSeasonQuarter(msg.getOrderCloseDate()));
        strDayIndex = DateUtils.getDay(new Date(msg.getOrderCloseDate()));
        dayIndex = Integer.parseInt(strDayIndex);
        cancelledOrderDaily.setDayIndex(dayIndex);
        cancelledOrderDaily.setSystemId(RptCommonUtils.getSystemId());
        if(msg.getCountyId()>0){
            cancelledOrderDaily.setCountyId(msg.getCountyId());
            Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
            RPTArea area = areaMap.get(msg.getCountyId());
            if(area != null){
                String[] split = area.getParentIds().split(",");
                if(split.length == 4){
                    cancelledOrderDaily.setCityId(Long.valueOf(split[3]));
                    cancelledOrderDaily.setProvinceId(Long.valueOf(split[2]));
                }
            }
        }
        cancelledOrderDaily.setServicePointId(msg.getServicePointId());
        try {
            cancelledOrderDailyRptMapper.insert(cancelledOrderDaily);
        }catch (Exception e){
            throw new RuntimeException("CancelledOrderDailyRptService.save保存退单取消单失败,失败原因:" + e.getMessage());
        }
    }



    /**
     * 根据订单结束时间时间和系统标识分页获取订单id和分片
     * @param startCloseDt
     * @param endCloseDt
     * @param quarter
     */
    public List<Long> findOrderIdByCloseDate(Long startCloseDt,Long endCloseDt,String quarter){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        Integer systemId = RptCommonUtils.getSystemId();
        List<Long> allCancelledOrderIdList = Lists.newArrayList();
        List<Long> cancelledOrderIdList = Lists.newArrayList();
        do{
            cancelledOrderIdList = cancelledOrderDailyRptMapper.findOrderIdByCloseDate(startCloseDt,endCloseDt,pageNum,pageSize,quarter,systemId);
            if(cancelledOrderIdList!=null && cancelledOrderIdList.size()>0){
                allCancelledOrderIdList.addAll(cancelledOrderIdList);
            }
            pageNum = pageNum + cancelledOrderIdList.size();
            listSize = cancelledOrderIdList.size();
        }while (listSize == pageSize);
        return allCancelledOrderIdList;
    }

    /**
     * 去掉退单取消单重复数据
     * @param date
     */
    public void deleteRepeatOrder(Date date){
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        Long startDt = startDate.getTime();
        Long endDt = endDate.getTime();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<Long> orderIdList = findOrderIdByCloseDate(startDt,endDt,quarter);
        if(orderIdList !=null && orderIdList.size()>0){
            Set<Long> set = new HashSet<>(orderIdList);
            //获得list与set的差集
            Collection rs = CollectionUtils.disjunction(orderIdList,set);
            //将collection转换为list
            List<Long> deleteList = new ArrayList<>(rs);
            if(deleteList!=null && deleteList.size()>0){
                Integer systemId = RptCommonUtils.getSystemId();
                for(Long orderId:deleteList){
                    Long id = cancelledOrderDailyRptMapper.getByOrderId(orderId,quarter,systemId);
                    if(id!=null && id>0){
                        cancelledOrderDailyRptMapper.deleteById(id,quarter);
                    }
                }
            }
        }
    }

    /**
     * 中间表重建增加
     * @param  date
     */
    public void saveCancelledOrderDailyRpt(Date date){
        int pageNum = 0;
        int pageSize = 5000;
        int listSize = 0;
        Date startDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<RPTCancelledOrderDailyEntity> cancelledOrderDailyEntities = Lists.newArrayList();
        Integer systemId = RptCommonUtils.getSystemId();
        do{
            cancelledOrderDailyEntities = cancelledOrderDailyRptMapper.getListFromWebDB(startDate,endDate,pageNum,pageSize);
            if(cancelledOrderDailyEntities !=null && cancelledOrderDailyEntities.size()>0){
                for(RPTCancelledOrderDailyEntity item:cancelledOrderDailyEntities){
                    if(item.getOrderCreateDate()!=null){
                        item.setOrderCreateDt(item.getOrderCreateDate().getTime());
                    }
                    if(item.getOrderCloseDate() !=null){
                        item.setOrderCloseDt(item.getOrderCloseDate().getTime());
                        item.setDayIndex(StringUtils.toInteger(DateUtils.getDay(item.getOrderCloseDate())));
                    }
                    item.setQuarter(quarter);
                    item.setSystemId(systemId);
                    cancelledOrderDailyRptMapper.insert(item);
                }
            }
            pageNum = pageNum + cancelledOrderDailyEntities.size();
            listSize = cancelledOrderDailyEntities.size();
        }while (listSize == pageSize);
    }


    /**
     * 补漏数据
     * @param  date
     */
    public void saveMissCancelledOrderRptToDB(Date date){
        int pageNum = 0;
        int pageSie = 5000;
        int listSize = 0;
        Date startCloseDate = DateUtils.getDateStart(date);
        Date endCloseDate = DateUtils.getDateEnd(date);
        Long startCloseDt = startCloseDate.getTime();
        Long endCloseDt = endCloseDate.getTime();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Integer systemId = RptCommonUtils.getSystemId();
        List<Long> cancelledOrderRptIdList = findOrderIdByCloseDate(startCloseDt,endCloseDt,quarter);
        List<Long> cancelledOrderIdListFromWebDB = Lists.newArrayList();
        List<RPTCancelledOrderDailyEntity> cancelledOrderListFromWebDB = Lists.newArrayList();
        Map<Long,RPTCancelledOrderDailyEntity> cancelledOrderMap = Maps.newHashMap();
        do{
            cancelledOrderListFromWebDB = cancelledOrderDailyRptMapper.getListFromWebDB(startCloseDate,endCloseDate,pageNum,pageSie);
            if(cancelledOrderListFromWebDB !=null && cancelledOrderListFromWebDB.size()>0){
                cancelledOrderIdListFromWebDB = cancelledOrderListFromWebDB.stream().map(RPTCancelledOrderDailyEntity::getOrderId).collect(Collectors.toList());
                cancelledOrderIdListFromWebDB.removeAll(cancelledOrderRptIdList);
                if(cancelledOrderIdListFromWebDB !=null && cancelledOrderIdListFromWebDB.size()>0){
                    cancelledOrderMap = cancelledOrderListFromWebDB.stream().collect(Collectors.toMap(RPTCancelledOrderDailyEntity::getOrderId,a->a,(k1,k2)->k1));
                    for(Long orderId:cancelledOrderIdListFromWebDB){
                        RPTCancelledOrderDailyEntity cancelledOrderRpt = cancelledOrderMap.get(orderId);
                        if(cancelledOrderRpt!=null){
                            if(cancelledOrderRpt.getOrderCreateDate() !=null){
                                cancelledOrderRpt.setOrderCreateDt(cancelledOrderRpt.getOrderCreateDate().getTime());
                            }
                            if(cancelledOrderRpt.getOrderCloseDate() !=null){
                                cancelledOrderRpt.setOrderCloseDt(cancelledOrderRpt.getOrderCloseDate().getTime());
                                cancelledOrderRpt.setDayIndex(StringUtils.toInteger(DateUtils.getDay(cancelledOrderRpt.getOrderCloseDate())));
                            }
                            cancelledOrderRpt.setSystemId(systemId);
                            cancelledOrderRpt.setQuarter(quarter);
                            cancelledOrderDailyRptMapper.insert(cancelledOrderRpt);
                        }
                    }
                }
            }
            pageNum = pageNum + cancelledOrderListFromWebDB.size();
            listSize = cancelledOrderListFromWebDB.size();
        }while (listSize == pageSie);
    }


    /**
     * 根据时间删除数据
     * @param  date
     */
     public void deleteByCloseDate(Date date){
         Date startDate = DateUtils.getDateStart(date);
         Date endDate = DateUtils.getDateEnd(date);
         String quarter = QuarterUtils.getSeasonQuarter(date);
         Integer systemId = RptCommonUtils.getSystemId();
         cancelledOrderDailyRptMapper.deleteByCloseDate(startDate.getTime(),endDate.getTime(),quarter,systemId);
     }



    /**
     *退单取消单中间数据重建
     *@param operationType
     *@param beginDt
     *@param endDt
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
                            saveCancelledOrderDailyRpt(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissCancelledOrderRptToDB(beginDate);
                            break;
                        case UPDATE:
                            deleteByCloseDate(beginDate);
                            saveCancelledOrderDailyRpt(beginDate);
                            break;
                        case DELETE:
                            deleteByCloseDate(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CancelledOrderDailyRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


}
