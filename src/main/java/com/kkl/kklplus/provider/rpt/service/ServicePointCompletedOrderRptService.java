package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTServicePointCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointCompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointCompletedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTEngineerPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTServicePointPbUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointCompletedOrderRptService {
    @Resource
    private ServicePointCompletedOrderRptMapper servicePointCompletedOrderRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    /**
     * 从Web数据库中取完工数据
     */
    private List<RPTServicePointCompletedOrderEntity> getServicePointCompletedOrderListFromWebDB(Date date) {
        List<RPTServicePointCompletedOrderEntity> result = Lists.newArrayList();
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<RPTServicePointCompletedOrderEntity> list = servicePointCompletedOrderRptMapper.getServicePointCompletedOrderListFromWebDB(beginDate, endDate);
        if (list != null && !list.isEmpty()) {
            List<RPTOrderDetail> orderDetailList = servicePointCompletedOrderRptMapper.getServicePointCompletedOrderDetailListFromWebDB(beginDate, endDate);
            List<RPTOrderServicePointFee> orderServicePointFeeList = servicePointCompletedOrderRptMapper.getOrderServicePointFeeListFromWebDB(beginDate, endDate);

            Map<Long, RPTServiceType> serviceTypeMap = MDUtils.getAllServiceTypeMap();
            Set<Long> productIds = orderDetailList.stream().map(i->i.getProduct().getId()).collect(Collectors.toSet());
            Map<Long, RPTProduct> productMap = MDUtils.getAllProductMap(Lists.newArrayList(productIds));
            RPTServiceType serviceType;
            RPTProduct product;
            RPTDict dataSourceDict, statusDict;

            Set<Long> servicePointIds = Sets.newHashSet();
            Set<Long> engineerIds = Sets.newHashSet();
            Map<Long, List<RPTOrderDetail>> allOrderDetailMap = Maps.newHashMap();
            for (RPTOrderDetail detail : orderDetailList) {
                servicePointIds.add(detail.getServicePointId());
                engineerIds.add(detail.getEngineerId());

                serviceType = serviceTypeMap.get(detail.getServiceTypeId());
                if (serviceType != null && StringUtils.isNotBlank(serviceType.getName())) {
                    detail.getServiceType().setName(serviceType.getName());
                }
                product = productMap.get(detail.getProductId());
                if (product != null && StringUtils.isNotBlank(product.getName())) {
                    detail.getProduct().setName(product.getName());
                }
                if (allOrderDetailMap.containsKey(detail.getOrderId())) {
                    allOrderDetailMap.get(detail.getOrderId()).add(detail);
                } else {
                    allOrderDetailMap.put(detail.getOrderId(), Lists.newArrayList(detail));
                }
            }
            Map<Long, List<RPTOrderServicePointFee>> allOrderServicePointFeeMap = orderServicePointFeeList.stream().collect(Collectors.groupingBy(RPTOrderServicePointFee::getOrderId));

            String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType"};
            Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));

            Map<String, RPTDict> dataSourceMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_DATA_SOURCE);
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
            Map<String, RPTDict> bankTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
            List<RPTOrderDetail> details;
            List<RPTOrderServicePointFee> pointFees;
            for (RPTServicePointCompletedOrderEntity item : list) {
                dataSourceDict = dataSourceMap.get(item.getDataSource().getValue());
                if (dataSourceDict != null && StringUtils.isNotBlank(dataSourceDict.getLabel())) {
                    item.getDataSource().setLabel(dataSourceDict.getLabel());
                }
                statusDict = statusMap.get(item.getStatus().getValue());
                if (statusDict != null && StringUtils.isNotBlank(statusDict.getLabel())) {
                    item.getStatus().setLabel(statusDict.getLabel());
                }
                details = allOrderDetailMap.get(item.getOrderId());
                List<RPTServicePoint> servicePoints = Lists.newArrayList();
                List<RPTEngineer> engineers = Lists.newArrayList();
                if (details != null && !details.isEmpty()) {
                    Set<Long> sIds = Sets.newHashSet();
                    Set<Long> eIds = Sets.newHashSet();
                    for (RPTOrderDetail detail : details) {
                        if (!sIds.contains(detail.getServicePointId())) {
                            MDServicePointViewModel servicePointVM = servicePointMap.get(detail.getServicePointId());
                            RPTServicePoint servicePoint = new RPTServicePoint();
                            if (servicePointVM != null) {
                                servicePoint.setServicePointNo(StringUtils.toString(servicePointVM.getServicePointNo()));
                                servicePoint.setName(StringUtils.toString(servicePointVM.getName()));
                                servicePoint.setContactInfo1(StringUtils.toString(servicePointVM.getContactInfo1()));
                                if (servicePointVM.getBank() != null) {
                                    RPTDict bankTypeDict = bankTypeMap.get(servicePointVM.getBank().toString());
                                    if (bankTypeDict != null) {
                                        servicePoint.setBank(bankTypeDict);
                                    }
                                }
                                servicePoint.setBankOwner(StringUtils.toString(servicePointVM.getBankOwner()));
                                servicePoint.setBankNo(StringUtils.toString(servicePointVM.getBankNo()));
                                if (servicePointVM.getPaymentType() != null) {
                                    RPTDict engineerPaymentTypeDict = paymentTypeMap.get(servicePointVM.getPaymentType().toString());
                                    if (engineerPaymentTypeDict != null) {
                                        servicePoint.setPaymentType(engineerPaymentTypeDict);
                                    }
                                }
                            }
                            servicePoint.setId(detail.getServicePointId());
                            servicePoints.add(servicePoint);
                            sIds.add(detail.getServicePointId());
                        }
                        if (!eIds.contains(detail.getEngineerId())) {
                            RPTEngineer engineer = engineerMap.get(detail.getEngineerId());
                            if (engineer == null) {
                                engineer = new RPTEngineer();
                            }
                            engineer.setId(detail.getEngineerId());
                            engineers.add(engineer);
                            eIds.add(detail.getEngineerId());
                        }
                    }
                    item.setDetails(details);
                    item.setServicePoints(servicePoints);
                    item.setEngineers(engineers);
                }
                pointFees = allOrderServicePointFeeMap.get(item.getOrderId());
                if (pointFees != null && !pointFees.isEmpty()) {
                    item.setOrderServicePointFees(pointFees);
                }
                result.add(item);
            }
        }
        return result;
    }

    private void insertServicePointCompletedOrder(RPTServicePointCompletedOrderEntity entity, String quarter, int systemId) {
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            entity.setAppointmentDt(entity.getAppointmentDate() != null ? entity.getAppointmentDate().getTime() : 0);
            entity.setCloseDt(entity.getCloseDate() != null ? entity.getCloseDate().getTime() : 0);//遇到了已对账的close_date时间为null的工单
            entity.setChargeDt(entity.getChargeDate().getTime());
            Map<Long, List<RPTOrderDetail>> detailMap = entity.getDetails().stream().collect(Collectors.groupingBy(RPTOrderDetail::getServicePointId));
            Map<Long, RPTServicePoint> servicePointMap = entity.getServicePoints().stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
            Map<Long, RPTEngineer> engineerMap = entity.getEngineers().stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
            Map<Long, RPTOrderServicePointFee> pointFeeMap = entity.getOrderServicePointFees().stream().collect(Collectors.toMap(RPTOrderServicePointFee::getServicePointId, i->i));

            double engineerTotalCharge;
            List<RPTOrderDetail> details;
            Set<Long> eIds;
            RPTServicePoint servicePoint;
            List<RPTEngineer> engineers;
            RPTOrderServicePointFee pointFee;
            RPTEngineer engineer;
            for (Long sid : detailMap.keySet()) {
                details = detailMap.get(sid);
                engineerTotalCharge = 0;
                eIds = Sets.newHashSet();
                engineers = Lists.newArrayList();
                for (RPTOrderDetail detail : details) {
                    engineerTotalCharge = engineerTotalCharge + detail.getEngineerTotalCharge();
                    if (!eIds.contains(detail.getEngineerId())) {
                        engineer = engineerMap.get(detail.getEngineerId());
                        if (engineer != null) {
                            engineers.add(engineer);
                        }
                    }
                    eIds.add(detail.getEngineerId());
                }

                servicePoint = servicePointMap.get(sid);
                if (servicePoint != null) {
                    entity.setServicePoint(servicePoint);
                    entity.setServicePointPb(RPTServicePointPbUtils.toServicePointBytes(servicePoint));
                }
                entity.setOrderDetailPb(RPTOrderDetailPbUtils.toOrderDetailsBytes(details));
                entity.setEngineerPb(RPTEngineerPbUtils.toEngineersBytes(engineers));
                entity.setEngineerSubtotalCharge(engineerTotalCharge);
                pointFee = pointFeeMap.get(sid);
                if (pointFee != null) {
                    entity.setEngineerInsuranceCharge(NumberUtils.toDouble(pointFee.getEngineerInsuranceCharge()));
                    entity.setEngineerTimelinessCharge(NumberUtils.toDouble(pointFee.getEngineerTimelinessCharge()));
                    entity.setEngineerCustomerTimelinessCharge(NumberUtils.toDouble(pointFee.getEngineerCustomerTimelinessCharge()));
                    entity.setEngineerUrgentCharge(NumberUtils.toDouble(pointFee.getEngineerUrgentCharge()));
                    entity.setEngineerPraiseFee(NumberUtils.toDouble(pointFee.getEngineerPraiseFee()));
                    entity.setEngineerTaxFee(NumberUtils.toDouble(pointFee.getEngineerTaxFee()));
                    entity.setEngineerInfoFee(NumberUtils.toDouble(pointFee.getEngineerInfoFee()));
                    entity.setEngineerDeposit(NumberUtils.toDouble(pointFee.getEngineerDeposit()));
                }
                servicePointCompletedOrderRptMapper.insert(entity);
            }
        } catch (Exception e) {
            log.error("【ServicePointCompletedOrderRptService.insertCompletedOrder】OrderId: {}, errorMsg: {}", entity.getOrderId(), Exceptions.getStackTraceAsString(e));
        }
    }

    /**
     * 将指定日期的网点完工单保存到中间表
     */
    public void saveServicePointCompletedOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointCompletedOrderEntity> list = getServicePointCompletedOrderListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                for (RPTServicePointCompletedOrderEntity item : list) {
                    if (!item.getDetails().isEmpty()) {
                        insertServicePointCompletedOrder(item, quarter, systemId);
                    }
                }
            }
        }
    }

    private Map<Long, Long> getServicePointCompletedOrderIdMap(Integer systemId, Date date) {
        Map<Long, Long> result = Maps.newHashMap();
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<LongTwoTuple> tuples = servicePointCompletedOrderRptMapper.getServicePointCompletedOrderIds(quarter, systemId, beginDate.getTime(), endDate.getTime());
        if (tuples != null && !tuples.isEmpty()) {
            for (LongTwoTuple item : tuples) {
                result.put(item.getBElement(), item.getAElement());
            }
        }
        return result;
    }

    /**
     * 将工单系统中有的而中间表中没有的网点完工单保存到中间表
     */
    public void saveMissedServicePointCompletedOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointCompletedOrderEntity> list = getServicePointCompletedOrderListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                Map<Long, Long> completedOrderIdMap = getServicePointCompletedOrderIdMap(systemId, date);
                Long primaryKeyId;
                for (RPTServicePointCompletedOrderEntity item : list) {
                    primaryKeyId = completedOrderIdMap.get(item.getOrderId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        insertServicePointCompletedOrder(item, quarter, systemId);
                    }
                }
            }
        }
    }

    /**
     * 删除中间表中指定日期的网点完工单明细
     */
    public void deleteServicePointCompletedOrdersFromRptDB(Date date) {
        if (date != null) {
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            servicePointCompletedOrderRptMapper.deleteServicePointCompletedOrders(quarter, systemId, beginDate.getTime(), endDate.getTime());
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
                            saveServicePointCompletedOrdersToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedServicePointCompletedOrdersToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteServicePointCompletedOrdersFromRptDB(beginDate);
                            saveServicePointCompletedOrdersToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointCompletedOrdersFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointCompletedOrderRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 分页查询网点完工单明细
     */
    public Page<RPTServicePointCompletedOrderEntity> getServicePointCompletedOrderListByPaging(RPTServicePointCompletedOrderSearch search) {
        Page<RPTServicePointCompletedOrderEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            search.setSystemId(RptCommonUtils.getSystemId());
            returnPage = servicePointCompletedOrderRptMapper.getServicePointCompletedOrderListByPaging(search);
            if (!returnPage.isEmpty()) {
                List<RPTOrderDetail> orderDetails;
                RPTServicePoint servicePoint;
                List<RPTEngineer> engineers;
                Map<Long, RPTEngineer> engineerMap;
                RPTEngineer engineer;
                for (RPTServicePointCompletedOrderEntity item : returnPage) {
                    if (item.getAppointmentDt() != 0) {
                        item.setAppointmentDate(new Date(item.getAppointmentDt()));
                    }
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setChargeDate(new Date(item.getChargeDt()));
                    orderDetails = RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
                    item.setDetails(orderDetails);
                    item.setDetails(null);
                    servicePoint = RPTServicePointPbUtils.fromServicePointBytes(item.getServicePointPb());
                    if (servicePoint != null) {
                        item.setServicePoint(servicePoint);
                        item.setServicePoints(Lists.newArrayList(servicePoint));
                    }
                    item.setServicePointPb(null);
                    engineers = RPTEngineerPbUtils.fromEngineersBytes(item.getEngineerPb());
                    item.setEngineers(engineers);
                    item.setEngineerPb(null);
                    engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                    for (RPTOrderDetail detail : orderDetails) {
                        if (servicePoint != null && servicePoint.getId() == detail.getServicePointId()) {
                            detail.setServicePoint(servicePoint);
                        }
                        engineer = engineerMap.get(detail.getEngineerId());
                        if (engineer != null) {
                            detail.setEngineer(engineer);
                        }
                    }
                }
            }
        }
        return returnPage;
    }


}
