package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTServicePointWriteOffEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointWriteOffRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTEngineerPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTServicePointPbUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.web.RPTOrderItemUtils;
import lombok.extern.slf4j.Slf4j;
import org.javatuples.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 网点退补单
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointWriteOffRptService extends RptBaseService {

    @Resource
    private ServicePointWriteOffRptMapper servicePointWriteOffRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private MSCustomerService msCustomerService;

    /**
     * 从Web数据库提取网点退补数据
     */
    private List<RPTServicePointWriteOffEntity> getServicePointWriteOffListFromWebDB(Date date) {
        List<RPTServicePointWriteOffEntity> result = Lists.newArrayList();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<RPTServicePointWriteOffEntity> list = servicePointWriteOffRptMapper.getServicePointWriteOffListFromWebDB(quarter, beginDate, endDate);
        List<RPTOrderDetail> allOrderDetailList = servicePointWriteOffRptMapper.getOrderDetailListFromWebDB(quarter, beginDate, endDate);

        Set<Long> missedOrderIds = Sets.newHashSet();
        Set<Long> userIds = Sets.newHashSet();
        Set<Long> customerIds = Sets.newHashSet();
        Set<Long> servicePointIds = Sets.newHashSet();
        Set<Long> engineerIds = Sets.newHashSet();
        for (RPTServicePointWriteOffEntity item : list) {
            if (StringUtils.isBlank(item.getOrderNo())) {
                missedOrderIds.add(item.getOrderId());
            } else {
                customerIds.add(item.getCustomer().getId());
                userIds.add(item.getCreateBy().getId());
                userIds.add(item.getKeFu().getId());
            }
            servicePointIds.add(item.getServicePoint().getId());
            engineerIds.add(item.getEngineer().getId());
        }

        List<List<Long>> missedOrderIdsList = Lists.partition(Lists.newArrayList(missedOrderIds), 20);
        List<RPTServicePointWriteOffEntity> missedOrderList = Lists.newArrayList();
        List<RPTServicePointWriteOffEntity> orders;
        List<RPTOrderDetail> details;
        for (List<Long> ids : missedOrderIdsList) {
            orders = servicePointWriteOffRptMapper.getOrderListByOrderIdsFromWebDB(ids);
            if (orders != null && !orders.isEmpty()) {
                for (RPTServicePointWriteOffEntity order : orders) {
                    customerIds.add(order.getCustomer().getId());
                    userIds.add(order.getCreateBy().getId());
                    userIds.add(order.getKeFu().getId());
                }
                missedOrderList.addAll(orders);
            }
            details = servicePointWriteOffRptMapper.getOrderDetailListByOrderIdsFromWebDB(ids);
            allOrderDetailList.addAll(details);
        }

        String[] fields = new String[]{"id", "name", "code", "salesId", "paymentType", "contractDate"};
        Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fields));
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        userIds.addAll(salesIds);

        Map<Long, List<RPTOrderDetail>> allOrderDetailMap = Maps.newHashMap();
        Map<Long, RPTServiceType> serviceTypeMap = MDUtils.getAllServiceTypeMap();
        Set<Long> productIds = allOrderDetailList.stream().map(i->i.getProduct().getId()).collect(Collectors.toSet());
        Map<Long, RPTProduct> productMap = MDUtils.getAllProductMap(Lists.newArrayList(productIds));
        RPTServiceType serviceType;
        RPTProduct product;
        Set<Long> detailIdSet = Sets.newHashSet();
        for (RPTOrderDetail item : allOrderDetailList) {
            if (!detailIdSet.contains(item.getId())) {
                detailIdSet.add(item.getId());
                serviceType = serviceTypeMap.get(item.getServiceTypeId());
                if (serviceType != null && StringUtils.isNotBlank(serviceType.getName())) {
                    item.getServiceType().setName(serviceType.getName());
                }
                product = productMap.get(item.getProductId());
                if (product != null && StringUtils.isNotBlank(product.getName())) {
                    item.getProduct().setName(product.getName());
                }
                if (allOrderDetailMap.containsKey(item.getOrderId())) {
                    allOrderDetailMap.get(item.getOrderId()).add(item);
                } else {
                    allOrderDetailMap.put(item.getOrderId(), Lists.newArrayList(item));
                }
            }
        }

        Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
        Map<String, RPTDict> bankTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);

        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));

        RPTCustomer customer;
        String userName;
        List<RPTOrderItem> items;
        RPTDict statusDict, paymentTypeDict;
        RPTServicePointWriteOffEntity missedOrder;
        MDServicePointViewModel servicePointVM;
        RPTEngineer engineer;
        List<RPTOrderItem> allOrderItemList = Lists.newArrayList();
        Map<Long, RPTServicePointWriteOffEntity> missedOrderMap = missedOrderList.stream().collect(Collectors.toMap(RPTServicePointWriteOffEntity::getOrderId, i -> i));
        for (RPTServicePointWriteOffEntity item : list) {
            Integer warrantyStatus = 10 ;
            if (StringUtils.isBlank(item.getOrderNo())) {
                missedOrder = missedOrderMap.get(item.getOrderId());
                if (missedOrder != null) {
                    item.setOrderNo(missedOrder.getOrderNo());
                    item.setWorkCardId(missedOrder.getWorkCardId());
                    item.setParentBizOrderId(missedOrder.getParentBizOrderId());
                    item.setDataSource(missedOrder.getDataSource());
                    item.setCustomer(missedOrder.getCustomer());
                    item.setPaymentType(missedOrder.getPaymentType());
                    item.setExpectCharge(missedOrder.getExpectCharge());
                    item.setBlockedCharge(missedOrder.getBlockedCharge());
                    item.setProductCategory(missedOrder.getProductCategory());
                    item.setShop(missedOrder.getShop());
                    item.setUserName(missedOrder.getUserName());
                    item.setUserPhone(missedOrder.getUserPhone());
                    item.setUserAddress(missedOrder.getUserAddress());
                    item.setStatus(missedOrder.getStatus());
                    item.setDescription(missedOrder.getDescription());
                    item.setCreateBy(missedOrder.getCreateBy());
                    item.setCreateDate(missedOrder.getCreateDate());
                    item.setKeFu(missedOrder.getKeFu());
                    item.setCloseDate(missedOrder.getCloseDate());
                    item.setChargeDate(missedOrder.getChargeDate());
                    item.setPlanDate(missedOrder.getPlanDate());
                    item.setAppointmentDate(missedOrder.getAppointmentDate());
//                    item.setOrderItemJson(missedOrder.getOrderItemJson());
                    item.setOrderItemPb(missedOrder.getOrderItemPb());
                }
            }
//            items = RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson());
            items = RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb());
            if (items != null && !items.isEmpty()) {
                item.setItems(items);
                allOrderItemList.addAll(items);
                for(RPTOrderItem orderItem : items){
                    if(orderItem.getServiceType()!= null && orderItem.getServiceType().getId()!= null){
                        serviceType = serviceTypeMap.get(orderItem.getServiceType().getId());
                        if(serviceType.getWarrantyStatus().equals(serviceType.WARRANTY_STATUS_OOT)){
                            warrantyStatus = 20;
                        }
                    }
                }
            }
            details = allOrderDetailMap.get(item.getOrderId());
            item.setWarrantyStatus(warrantyStatus);
            if (details != null && !details.isEmpty()) {
                item.setDetails(details);
            }
            statusDict = statusMap.get(item.getStatus().getValue());
            if (statusDict != null && StringUtils.isNotBlank(statusDict.getLabel())) {
                item.getStatus().setLabel(statusDict.getLabel());
            }
            customer = customerMap.get(item.getCustomer().getId());
            if (customer != null) {
                if (StringUtils.isNotBlank(customer.getName())) {
                    item.getCustomer().setName(customer.getName());
                }
                if (StringUtils.isNotBlank(customer.getCode())) {
                    item.getCustomer().setCode(customer.getCode());
                }
                if (customer.getContractDate() != null) {
                    item.getCustomer().setContractDate(customer.getContractDate());
                    item.getCustomer().setContractDt(customer.getContractDate().getTime());
                }
                if (customer.getPaymentType() != null && StringUtils.isNotBlank(customer.getPaymentType().getValue())) {
                    paymentTypeDict = paymentTypeMap.get(customer.getPaymentType().getValue());
                    if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                        item.getCustomer().getPaymentType().setLabel(paymentTypeDict.getLabel());
                    }
                }
                if (customer.getSales() != null && customer.getSales().getId() != null) {
                    item.getCustomer().getSales().setId(customer.getSales().getId());
                    userName = userNameMap.get(customer.getSales().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCustomer().getSales().setName(userName);
                    }
                }
            }
            userName = userNameMap.get(item.getCreateBy().getId());
            if (StringUtils.isNotBlank(userName)) {
                item.getCreateBy().setName(userName);
            }
            if (item.getKeFu() != null && item.getKeFu().getId() != null) {
                userName = userNameMap.get(item.getKeFu().getId());
                if (StringUtils.isNotBlank(userName)) {
                    item.getKeFu().setName(userName);
                }
            }
            int dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
            if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                item.getShop().setLabel(shopName);
            }
            servicePointVM = servicePointMap.get(item.getServicePoint().getId());
            if (servicePointVM != null) {
                item.getServicePoint().setServicePointNo(StringUtils.toString(servicePointVM.getServicePointNo()));
                item.getServicePoint().setName(StringUtils.toString(servicePointVM.getName()));
                item.getServicePoint().setContactInfo1(StringUtils.toString(servicePointVM.getContactInfo1()));
                if (servicePointVM.getBank() != null) {
                    RPTDict bankTypeDict = bankTypeMap.get(servicePointVM.getBank().toString());
                    if (bankTypeDict != null) {
                        item.getServicePoint().setBank(bankTypeDict);
                    }
                }
                item.getServicePoint().setBankOwner(StringUtils.toString(servicePointVM.getBankOwner()));
                item.getServicePoint().setBankNo(StringUtils.toString(servicePointVM.getBankNo()));
                if (servicePointVM.getPaymentType() != null) {
                    RPTDict engineerPaymentTypeDict = paymentTypeMap.get(servicePointVM.getPaymentType().toString());
                    if (engineerPaymentTypeDict != null) {
                        item.getServicePoint().setPaymentType(engineerPaymentTypeDict);
                    }
                }
            }
            engineer = engineerMap.get(item.getEngineer().getId());
            if (engineer != null) {
                item.getEngineer().setName(engineer.getName());
            }
            result.add(item);
        }
        RPTOrderItemUtils.setOrderItemProperties(allOrderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
        return result;
    }

    private void insertServicePointWriteOff(RPTServicePointWriteOffEntity entity, String quarter, int systemId) {
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            entity.setCreateDt(entity.getCreateDate().getTime());
            entity.setCloseDt(entity.getCloseDate().getTime());
            entity.setChargeDt(entity.getChargeDate().getTime());
            entity.setPlanDt(entity.getPlanDate() != null ? entity.getPlanDate().getTime() : 0);
            entity.setAppointmentDt(entity.getAppointmentDate() != null ? entity.getAppointmentDate().getTime() : 0);
            entity.setWriteOffCreateDt(entity.getWriteOffCreateDate().getTime());
            entity.setOrderItemPb(RPTOrderItemPbUtils.toOrderItemsBytes(entity.getItems()));
            entity.setOrderDetailPb(RPTOrderDetailPbUtils.toOrderDetailsBytes(entity.getDetails()));
            entity.setServicePointPb(RPTServicePointPbUtils.toServicePointBytes(entity.getServicePoint()));
            entity.setEngineerPb(RPTEngineerPbUtils.toEngineerBytes(entity.getEngineer()));
            servicePointWriteOffRptMapper.insert(entity);
        } catch (Exception e) {
            log.error("【ServicePointWriteOffRptService.insertServicePointWriteOff】ServicePointWriteOffId: {}, errorMsg: {}", entity.getServicePointWriteOffId(), Exceptions.getStackTraceAsString(e));
        }

    }

    private Map<Long, Long> getServicePointWriteOffIdMap(Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<LongTwoTuple> tuples = servicePointWriteOffRptMapper.getServicePointWriteOffIds(quarter, RptCommonUtils.getSystemId(), beginDate.getTime(), endDate.getTime());
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

    /**
     * 将指定日期的退补单保存到中间表
     */
    public void saveServicePointWriteOffToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointWriteOffEntity> list = getServicePointWriteOffListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                for (RPTServicePointWriteOffEntity item : list) {
                    insertServicePointWriteOff(item, quarter, systemId);
                }
            }
        }
    }

    /**
     * 将工单系统中有的而中间表中没有的客户退补保存到中间表
     */
    public void saveMissedServicePointWriteOffsToRptDB(Date date) {
        if (date != null) {
            List<RPTServicePointWriteOffEntity> list = getServicePointWriteOffListFromWebDB(date);
            if (!list.isEmpty()) {
                Map<Long, Long> servicePointWriteOffIdMap = getServicePointWriteOffIdMap(date);
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                Long primaryKeyId;
                for (RPTServicePointWriteOffEntity item : list) {
                    primaryKeyId = servicePointWriteOffIdMap.get(item.getServicePointWriteOffId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        insertServicePointWriteOff(item, quarter, systemId);
                    }
                }
            }
        }
    }

    /**
     * 删除中间表中指定日期的退单或取消单明细
     */
    public void deleteServicePointWriteOffsFromRptDB(Date date) {
        if (date != null) {
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            servicePointWriteOffRptMapper.deleteServicePointWriteOffs(quarter, systemId, beginDate.getTime(), endDate.getTime());
        }
    }

    /**
     * 分页查询退单/取消单明细
     */
    public Page<RPTServicePointWriteOffEntity> getServicePointWriteOffListByPaging(RPTServicePointWriteOffSearch search) {
        Page<RPTServicePointWriteOffEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            search.setSystemId(RptCommonUtils.getSystemId());
            returnPage = servicePointWriteOffRptMapper.getServicePointWriteOffListByPaging(search);
            if (!returnPage.isEmpty()) {
                RPTServicePoint servicePoint;
                RPTEngineer engineer;
                for (RPTServicePointWriteOffEntity item : returnPage) {
                    item.setCreateDate(new Date(item.getCreateDt()));
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setChargeDate(new Date(item.getChargeDt()));
                    if (item.getPlanDt() != 0) {
                        item.setPlanDate(new Date(item.getPlanDt()));
                    }
                    if (item.getAppointmentDt() != 0) {
                        item.setAppointmentDate(new Date(item.getAppointmentDt()));
                    }
                    item.setWriteOffCreateDate(new Date(item.getWriteOffCreateDt()));
                    item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                    item.setOrderItemPb(null);
                    item.setDetails(RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb()));
                    item.setOrderDetailPb(null);
                    servicePoint = RPTServicePointPbUtils.fromServicePointBytes(item.getServicePointPb());
                    if (servicePoint != null) {
                        item.setServicePoint(servicePoint);
                    }
                    engineer = RPTEngineerPbUtils.fromEngineerBytes(item.getEngineerPb());
                    if (engineer != null) {
                        item.setEngineer(engineer);
                    }
                }
            }
        }
        return returnPage;
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
                            saveServicePointWriteOffToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedServicePointWriteOffsToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteServicePointWriteOffsFromRptDB(beginDate);
                            saveServicePointWriteOffToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointWriteOffsFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointWriteOffRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

}
