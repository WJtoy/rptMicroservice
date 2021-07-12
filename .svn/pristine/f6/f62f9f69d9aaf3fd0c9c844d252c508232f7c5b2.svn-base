package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDActionCode;
import com.kkl.kklplus.entity.md.MDErrorCode;
import com.kkl.kklplus.entity.md.MDErrorType;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.common.WarrantyStatusEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CompletedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSErrorService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTEngineerPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTServicePointPbUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.provider.rpt.utils.web.RPTOrderItemUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.javatuples.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CompletedOrderRptService extends RptBaseService {
    @Resource
    private CompletedOrderRptMapper completedOrderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private MSErrorService msErrorService;

    /**
     * 从Web数据库中取完工数据
     */
    private List<RPTCompletedOrderEntity> getCompletedOrderListFromWebDB(Date date) {
        List<RPTCompletedOrderEntity> result = Lists.newArrayList();
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<RPTCompletedOrderEntity> list = completedOrderRptMapper.getCompletedOrderListFromWebDB(beginDate, endDate);
        if (list != null && !list.isEmpty()) {
            List<RPTOrderDetail> orderDetailList = completedOrderRptMapper.getCompletedOrderDetailListFromWebDB(beginDate, endDate);
            List<RPTOrderItem> orderItemList = completedOrderRptMapper.getCompletedOrderUnitBarcodeListFromWebDB(beginDate, endDate);

            List<RPTOrderDetail> orderDetailFees = completedOrderRptMapper.getCompletedOrderDetailFeeListWebDB(beginDate, endDate);
            Map<Long, List<RPTOrderDetail>> orderDetailFeeMap = orderDetailFees.stream().collect(Collectors.groupingBy(RPTOrderDetail::getOrderId));

            Map<String, List<String>> unitBarcodeMap = Maps.newHashMap();
            if (orderItemList != null && !orderItemList.isEmpty()) {
                String key;
                for (RPTOrderItem item : orderItemList) {
                    key = String.format("%d:%d", item.getOrderId(), item.getProductId());
                    if (unitBarcodeMap.containsKey(key)) {
                        if (StringUtils.isNotBlank(item.getUnitBarcode())) {
                            unitBarcodeMap.get(key).add(item.getUnitBarcode());
                        }
                    } else {
                        if (StringUtils.isNotBlank(item.getUnitBarcode())) {
                            unitBarcodeMap.put(key, Lists.newArrayList(item.getUnitBarcode()));
                        }
                    }
                }
            }

            Set<Long> customerIds = Sets.newHashSet();
            Set<Long> servicePointIds = Sets.newHashSet();
            Set<Long> engineerIds = Sets.newHashSet();
            Set<Long> userIds = Sets.newHashSet();

            Map<String, RPTDict> dataSourceMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_DATA_SOURCE);
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
            Map<String, RPTDict> appCompleteTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_APP_COMPLETE_TYPE);
            Map<String, RPTDict> bankTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
            Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
            Map<Long, RPTServiceType> serviceTypeMap = MDUtils.getAllServiceTypeMap();

            RPTServiceType serviceType;
            RPTProduct product;
            RPTCustomer customer;
            String userName;
            RPTDict dataSourceDict, paymentTypeDict, statusDict, appCompleteTypeDict;
            Set<Long> productIds = orderDetailList.stream().map(i -> i.getProduct().getId()).collect(Collectors.toSet());
            Map<Long, RPTProduct> productMap = MDUtils.getAllProductMap(Lists.newArrayList(productIds));
            List<RPTOrderItem> allOrderItems = Lists.newArrayList();
            for (RPTCompletedOrderEntity item : list) {
                Integer warrantyStatus = 10;
                customerIds.add(item.getCustomer().getId());
                userIds.add(item.getCreateBy().getId());
                userIds.add(item.getKeFu().getId());

                dataSourceDict = dataSourceMap.get(item.getDataSource().getValue());
                if (dataSourceDict != null && StringUtils.isNotBlank(dataSourceDict.getLabel())) {
                    item.getDataSource().setLabel(dataSourceDict.getLabel());
                }
                paymentTypeDict = paymentTypeMap.get(item.getPaymentType().getValue());
                if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                    item.getPaymentType().setLabel(paymentTypeDict.getLabel());
                }
                statusDict = statusMap.get(item.getStatus().getValue());
                if (statusDict != null && StringUtils.isNotBlank(statusDict.getLabel())) {
                    item.getStatus().setLabel(statusDict.getLabel());
                }
                int dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
                if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                    String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                    item.getShop().setLabel(shopName);
                }
                if (item.getAppCompleteType() != null && StringUtils.isNotBlank(item.getAppCompleteType().getValue())) {
                    appCompleteTypeDict = appCompleteTypeMap.get(item.getAppCompleteType().getValue());
                    if (appCompleteTypeDict != null && StringUtils.isNotBlank(appCompleteTypeDict.getLabel())) {
                        item.getAppCompleteType().setLabel(appCompleteTypeDict.getLabel());
                    }
                }
//                item.setItems(RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson()));
                item.setItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb()));
                for (RPTOrderItem orderItem : item.getItems()) {
                    String key = String.format("%d:%d", item.getOrderId(), orderItem.getProductId());
                    List<String> unitBarCodes = unitBarcodeMap.get(key);
                    if (unitBarCodes != null && !unitBarCodes.isEmpty()) {
                        orderItem.setUnitBarCodes(unitBarCodes);
                    }
                    if(orderItem.getServiceType()!= null && orderItem.getServiceType().getId()!= null){
                        serviceType = serviceTypeMap.get(orderItem.getServiceType().getId());
                        if(serviceType.getWarrantyStatus().equals(serviceType.WARRANTY_STATUS_OOT)){
                            warrantyStatus = 20;
                        }
                    }

                }
                allOrderItems.addAll(item.getItems());
                item.setWarrantyStatus(warrantyStatus);
            }
            RPTOrderItemUtils.setOrderItemProperties(allOrderItems, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));


            String[] fields = new String[]{"id", "name", "code", "salesId", "paymentType", "contractDate"};
            Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fields));
            Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                    .map(i -> i.getSales().getId()).collect(Collectors.toSet());
            userIds.addAll(salesIds);
            Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));


            Set<Long> errorTypeIds = new HashSet<>();
            Set<Long> errorCodeIds = new HashSet<>();
            Set<Long> actionCodeIds = new HashSet<>();

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
                errorTypeIds.add(Long.valueOf(detail.getErrorType().getValue()));
                errorCodeIds.add(Long.valueOf(detail.getErrorCode().getValue()));
                actionCodeIds.add(Long.valueOf(detail.getActionCode().getValue()));
            }

            String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType"};
            Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));

            Map<Long, MDErrorType> mdErrorTypeMap = msErrorService.findListByErrorTypeNameToMap(Lists.newArrayList(errorTypeIds));

            Map<Long, MDErrorCode> mdErrorCodeMap = msErrorService.findListByErrorCodeNameToMap(Lists.newArrayList(errorCodeIds));

            Map<Long, MDActionCode> mdActionCodeMap = msErrorService.findListByActionCodeNameToMap(Lists.newArrayList(actionCodeIds));
            List<RPTOrderDetail> details;
            List<RPTOrderDetail> detailFees;
            long servicePointId;
            int time;
            for (RPTCompletedOrderEntity item : list) {
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

                details = allOrderDetailMap.get(item.getOrderId());
                detailFees = orderDetailFeeMap.get(item.getOrderId());
                List<RPTServicePoint> servicePoints = Lists.newArrayList();
                List<RPTEngineer> engineers = Lists.newArrayList();
                if (details != null && !details.isEmpty()) {
                    details = details.stream().sorted(Comparator.comparing(RPTOrderDetail::getServicePointId).reversed()).collect(Collectors.toList());
                    Set<Long> sIds = Sets.newHashSet();
                    Set<Long> eIds = Sets.newHashSet();
                    if(detailFees != null && !detailFees.isEmpty()){
                        for (RPTOrderDetail detailFee : detailFees) {
                            time = 0;
                            for (RPTOrderDetail detail : details.stream().filter(p->p.getServicePointId() == detailFee.getServicePointId()).collect(Collectors.toList())) {
                                if (time < 1){
                                    servicePointId = detailFee.getServicePointId();
                                    if (detail.getServicePointId() == servicePointId) {
                                        detail.setEngineerInsuranceCharge(detailFee.getEngineerInsuranceCharge());
                                        detail.setEngineerTimelinessCharge(detailFee.getEngineerTimelinessCharge());
                                        detail.setEngineerCustomerTimelinessCharge(detailFee.getEngineerCustomerTimelinessCharge());
                                        detail.setEngineerPraiseFee(detailFee.getEngineerPraiseFee());
                                        detail.setEngineerTaxFee(detailFee.getEngineerTaxFee());
                                        detail.setEngineerInfoFee(detailFee.getEngineerInfoFee());
                                        detail.setEngineerDeposit(detailFee.getEngineerDeposit());
                                        detail.setEngineerUrgentCharge(detailFee.getEngineerUrgentCharge());
                                    }
                                    time++;
                                }

                            }
                        }
                    }

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

                        if (mdErrorTypeMap.get(Long.valueOf(detail.getErrorType().getValue())) != null) {
                            detail.getErrorType().setLabel(mdErrorTypeMap.get(Long.valueOf(detail.getErrorType().getValue())).getName());
                        }
                        if (mdErrorCodeMap.get(Long.valueOf(detail.getErrorCode().getValue())) != null) {
                            detail.getErrorCode().setLabel(mdErrorCodeMap.get(Long.valueOf(detail.getErrorCode().getValue())).getName());
                        }
                        if (mdActionCodeMap.get(Long.valueOf(detail.getActionCode().getValue())) != null) {
                            detail.getActionCode().setLabel(mdActionCodeMap.get(Long.valueOf(detail.getActionCode().getValue())).getName());
                        } else if (detail.getActionCode().getValue().equals("0")) {
                            detail.getActionCode().setLabel(detail.getActionName());
                        }

                    }


                    item.setDetails(details);
                    item.setServicePoints(servicePoints);
                    item.setEngineers(engineers);
                }
                result.add(item);
            }
        }
        return result;
    }

    private void insertCompletedOrder(RPTCompletedOrderEntity entity, String quarter, int systemId) {
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            entity.setCreateDt(entity.getCreateDate().getTime());
            entity.setCustomerApproveDt(entity.getCustomerApproveDate() == null ? 0 : entity.getCustomerApproveDate().getTime());
            entity.setPlanDt(entity.getPlanDate() != null ? entity.getPlanDate().getTime() : 0);
            entity.setArrivalDt(entity.getArrivalDate() != null ? entity.getArrivalDate().getTime() : 0);
            entity.setAppCompleteDt(entity.getAppCompleteDate() != null ? entity.getAppCompleteDate().getTime() : 0);
            entity.setAppointmentDt(entity.getAppointmentDate() != null ? entity.getAppointmentDate().getTime() : 0);
            entity.setCloseDt(entity.getCloseDate() != null ? entity.getCloseDate().getTime() : 0); //TODO: 系统中存在已对账且closeDate为NULL的工单
            entity.setChargeDt(entity.getChargeDate().getTime());
            entity.setOrderItemPb(RPTOrderItemPbUtils.toOrderItemsBytes(entity.getItems()));
            entity.setOrderDetailPb(RPTOrderDetailPbUtils.toOrderDetailsBytes(entity.getDetails()));
            entity.setEngineerSubtotalCharge(RPTOrderDetail.calcEngineerTotalCharge(entity.getDetails()));
            entity.setCustomerSubtotalCharge(RPTOrderDetail.calcCustomerTotalCharge(entity.getDetails()));
            entity.setServicePointPb(RPTServicePointPbUtils.toServicePointsBytes(entity.getServicePoints()));
            entity.setEngineerPb(RPTEngineerPbUtils.toEngineersBytes(entity.getEngineers()));
            completedOrderRptMapper.insert(entity);
        } catch (Exception e) {
            log.error("【CompletedOrderRptService.insertCompletedOrder】OrderId: {}, errorMsg: {}", entity.getOrderId(), Exceptions.getStackTraceAsString(e));
        }
    }

    private Map<Long, Long> getCompletedOrderIdMap(Integer systemId, Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<LongTwoTuple> tuples = completedOrderRptMapper.getCompletedOrderIds(quarter, systemId, beginDate.getTime(), endDate.getTime());
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

    /**
     * 将指定日期的完工单保存到中间表
     */
    public void saveCompletedOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTCompletedOrderEntity> list = getCompletedOrderListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                for (RPTCompletedOrderEntity item : list) {
                    insertCompletedOrder(item, quarter, systemId);
                }
            }
        }
    }

    /**
     * 将工单系统中有的而中间表中没有的完工单保存到中间表
     */
    public void saveMissedCompletedOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTCompletedOrderEntity> list = getCompletedOrderListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                Map<Long, Long> completedOrderIdMap = getCompletedOrderIdMap(systemId, date);
                Long primaryKeyId;
                for (RPTCompletedOrderEntity item : list) {
                    primaryKeyId = completedOrderIdMap.get(item.getOrderId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        insertCompletedOrder(item, quarter, systemId);
                    }
                }
            }
        }
    }

    /**
     * 删除中间表中指定日期的完工单明细
     */
    public void deleteCompletedOrdersFromRptDB(Date date) {
        if (date != null) {
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            completedOrderRptMapper.deleteCompletedOrders(quarter, systemId, beginDate.getTime(), endDate.getTime());
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
                            saveCompletedOrdersToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedCompletedOrdersToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteCompletedOrdersFromRptDB(beginDate);
                            saveCompletedOrdersToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteCompletedOrdersFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CompletedOrderRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 分页查询完工单明细
     */
    public Page<RPTCompletedOrderEntity> getCompletedOrderListByPaging(RPTCompletedOrderSearch search) {
        Page<RPTCompletedOrderEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            search.setSystemId(RptCommonUtils.getSystemId());
            returnPage = completedOrderRptMapper.getCompletedOrderListByPaging(search);
            if (!returnPage.isEmpty()) {
                List<RPTOrderDetail> orderDetails;
                List<RPTServicePoint> servicePoints;
                List<RPTEngineer> engineers;
                Map<Long, RPTServicePoint> servicePointMap;
                Map<Long, RPTEngineer> engineerMap;
                RPTServicePoint servicePoint;
                RPTEngineer engineer;
                for (RPTCompletedOrderEntity item : returnPage) {
                    item.setCreateDate(new Date(item.getCreateDt()));

                    if (item.getCustomerApproveDt() != 0) {
                        item.setCustomerApproveDate(new Date(item.getCustomerApproveDt()));
                    }
                    if (item.getPlanDt() != 0) {
                        item.setPlanDate(new Date(item.getPlanDt()));
                    }
                    if (item.getArrivalDt() != 0) {
                        item.setArrivalDate(new Date(item.getArrivalDt()));
                    }
                    if (item.getAppointmentDt() != 0) {
                        item.setAppointmentDate(new Date(item.getAppointmentDt()));
                    }
                    if (item.getAppCompleteDt() != 0) {
                        item.setAppCompleteDate(new Date(item.getAppCompleteDt()));
                    }
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setChargeDate(new Date(item.getChargeDt()));
                    item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                    item.setOrderItemPb(null);
                    orderDetails = RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
                    item.setDetails(orderDetails);
                    item.setOrderDetailPb(null);
                    servicePoints = RPTServicePointPbUtils.fromServicePointsBytes(item.getServicePointPb());
                    item.setServicePoints(servicePoints);
                    item.setServicePointPb(null);
                    engineers = RPTEngineerPbUtils.fromEngineersBytes(item.getEngineerPb());
                    item.setEngineers(engineers);
                    item.setEngineerPb(null);
                    servicePointMap = servicePoints.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                    engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                    for (RPTOrderDetail detail : orderDetails) {
                        servicePoint = servicePointMap.get(detail.getServicePointId());
                        if (servicePoint != null) {
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

    /**
     * 查询完工单明细
     */
    public List<RPTCompletedOrderEntity> getCompletedOrderList(RPTCustomerChargeSearch search) {
        List<RPTCompletedOrderEntity> returnList = null;
        if (search != null && search.getCustomerId() != null && search.getCustomerId() > 0
                && search.getSelectedYear() != null && search.getSelectedYear() > 0
                && search.getSelectedMonth() != null && search.getSelectedMonth() > 0) {

            Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
            String quarter = QuarterUtils.getSeasonQuarter(queryDate);
            Date beginDate = DateUtils.getStartOfDay(queryDate);
            Date endDate = DateUtils.addMonth(beginDate, 1);
            long beginDateTime = beginDate.getTime();
            long endDateTime = endDate.getTime();
            int systemId = RptCommonUtils.getSystemId();
            returnList = completedOrderRptMapper.getCompletedOrderList(systemId, beginDateTime, endDateTime, search.getCustomerId(), quarter);
            if (!returnList.isEmpty()) {
                List<RPTOrderDetail> orderDetails;
                List<RPTServicePoint> servicePoints;
                List<RPTEngineer> engineers;
                Map<Long, RPTServicePoint> servicePointMap;
                Map<Long, RPTEngineer> engineerMap;
                RPTServicePoint servicePoint;
                RPTEngineer engineer;
                for (RPTCompletedOrderEntity item : returnList) {
                    item.setCreateDate(new Date(item.getCreateDt()));
                    item.setCustomerApproveDate(new Date(item.getCustomerApproveDt()));
                    if (item.getCustomerApproveDt() != 0) {
                        item.setCustomerApproveDate(new Date(item.getCustomerApproveDt()));
                    }
                    if (item.getPlanDt() != 0) {
                        item.setPlanDate(new Date(item.getPlanDt()));
                    }
                    if (item.getAppointmentDt() != 0) {
                        item.setAppointmentDate(new Date(item.getAppointmentDt()));
                    }
                    if (item.getArrivalDt() != 0) {
                        item.setArrivalDate(new Date(item.getArrivalDt()));
                    }
                    if (item.getAppCompleteDt() != 0) {
                        item.setAppCompleteDate(new Date(item.getAppCompleteDt()));
                    }
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setChargeDate(new Date(item.getChargeDt()));
                    item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                    item.setOrderItemPb(null);
                    item.setOrderItemJson(null);
                    orderDetails = RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
                    item.setDetails(orderDetails);
                    item.setOrderDetailPb(null);
                    servicePoints = RPTServicePointPbUtils.fromServicePointsBytes(item.getServicePointPb());
                    item.setServicePoints(servicePoints);
                    item.setServicePointPb(null);
                    engineers = RPTEngineerPbUtils.fromEngineersBytes(item.getEngineerPb());
                    item.setEngineers(engineers);
                    item.setEngineerPb(null);
                    servicePointMap = servicePoints.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                    engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                    for (RPTOrderDetail detail : orderDetails) {
                        servicePoint = servicePointMap.get(detail.getServicePointId());
                        if (servicePoint != null) {
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
        return returnList;
    }


    /**
     * 添加完工单Sheet
     *
     * @param xBook
     * @param xStyle
     * @param orderMasterList
     */
    public void addCustomerChargeCompleteRptSheet(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCompletedOrderEntity> orderMasterList) {

        String xName = "完工单";
        Sheet xSheet = xBook.createSheet(xName);
        xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
        int rowIndex = 0;

        //====================================================绘制标题行============================================================
        Row titleRow = xSheet.createRow(rowIndex++);
        titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
        ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 41));

        //====================================================绘制表头============================================================
        //表头第一行
        Row headerFirstRow = xSheet.createRow(rowIndex++);
        headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 3));

        ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 4, 16));

        ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

        ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

        ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "APP完成时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

        ExportExcel.createCell(headerFirstRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 20, 20));

        ExportExcel.createCell(headerFirstRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 21, 26));

        ExportExcel.createCell(headerFirstRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 27, 36));

        ExportExcel.createCell(headerFirstRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 37, 37));

        ExportExcel.createCell(headerFirstRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工结果");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 38, 38));
        ExportExcel.createCell(headerFirstRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 39, 41));
        ExportExcel.createCell(headerFirstRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结账日期");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 42, 42));

        //表头第二行
        Row headerSecondRow = xSheet.createRow(rowIndex++);
        headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);


        ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
        ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
        ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");

        ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
        ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
        ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工条码");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

        ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
        ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");

        ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
        ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
        ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
        ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
        ExportExcel.createCell(headerSecondRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
        ExportExcel.createCell(headerSecondRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
        ExportExcel.createCell(headerSecondRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
        ExportExcel.createCell(headerSecondRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
        ExportExcel.createCell(headerSecondRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
        ExportExcel.createCell(headerSecondRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收合计");

        ExportExcel.createCell(headerSecondRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障类型");
        ExportExcel.createCell(headerSecondRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障现象");
        ExportExcel.createCell(headerSecondRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障处理");

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

        //====================================================绘制表格数据单元格============================================================
        int totalQty = 0;
        int totalActualQty = 0;
        double totalExpectCharge = 0.0;
        double totalCharge = 0.0;
        double totalDetailCharge = 0.0;
        double totalTimelinessCharge = 0.0;
        double totalUrgentCharge = 0.0;
        double totalCustomerPraiseFee = 0.0;
        if (orderMasterList != null) {
            for (int i = 0; i < orderMasterList.size(); i++) {

                RPTCompletedOrderEntity orderMaster = orderMasterList.get(i);
                int rowSpan = orderMaster.getMaxRow() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();
                List<RPTOrderDetail> detailList = orderMaster.getDetails();

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                }


                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                }
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                }


                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                }

                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                }
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                }
                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem item = itemList.get(0);
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.joinStringList(item.getUnitBarCodes()));
                    totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                totalExpectCharge = totalExpectCharge + expectCharge;
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                }

                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                }

                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                }

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                }

                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                }

                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                }
                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getArrivalDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                }
                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getAppCompleteDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                }
                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 20, 20));
                }

                RPTOrderDetail detail = null;
                if (detailList != null && detailList.size() > 0) {
                    detail = detailList.get(0);
                }

                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));
                int qty = (detail == null ? 0 : detail.getQty());
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                totalActualQty = totalActualQty + qty;

                double serviceCharge = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge);

                double materialCharge = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, materialCharge);

                double travelCharge = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, travelCharge);

                double expressCharge = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expressCharge);

                double otherCharge = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, otherCharge);
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge + materialCharge + travelCharge + expressCharge + otherCharge);
                totalDetailCharge = totalDetailCharge + serviceCharge + materialCharge + travelCharge + expressCharge + otherCharge;
                double customerTimeLinessCharge = orderMaster.getCustomerTimelinessCharge();
                totalTimelinessCharge = totalTimelinessCharge + customerTimeLinessCharge;
                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimeLinessCharge);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 33, 33));
                }
                double customerUrgentCharge = orderMaster.getCustomerUrgentCharge();
                totalUrgentCharge = totalUrgentCharge + customerUrgentCharge;
                ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerUrgentCharge);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 34, 34));
                }
                double customerPraiseFeeCharge = orderMaster.getCustomerPraiseFee();
                totalCustomerPraiseFee = totalCustomerPraiseFee + customerPraiseFeeCharge;
                ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerPraiseFeeCharge);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 35, 35));
                }
                ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTotalCharge());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 36, 36));
                }
                totalCharge = (totalCharge + orderMaster.getCustomerTotalCharge());

                ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "完成");
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 37, 37));
                }
                ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getAppCompleteType() == null ? "" : orderMaster.getAppCompleteType().getLabel());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 38, 38));
                }
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorType() == null ? "" : detail.getErrorType().getLabel());


                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorCode() == null ? "" : detail.getErrorCode().getLabel());

                ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getActionCode() == null ? "" : detail.getActionCode().getLabel());

                ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getChargeDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 42, 42));
                }


                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem item = itemList.get(index);

                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.joinStringList(item.getUnitBarCodes()));
                            totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        } else {
                            for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                        if (detailList != null && index < detailList.size()) {
                            detail = detailList.get(index);

                            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));
                            qty = (detail == null ? 0 : detail.getQty());
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                            totalActualQty = totalActualQty + qty;

                            double serviceCharge1 = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge1);

                            double materialCharge1 = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, materialCharge1);

                            double travelCharge1 = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, travelCharge1);

                            double expressCharge1 = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expressCharge1);

                            double otherCharge1 = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, otherCharge1);

                            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge1 + materialCharge1 + travelCharge1 + expressCharge1 + otherCharge1);
                            totalDetailCharge = totalDetailCharge + serviceCharge1 + materialCharge1 + travelCharge1 + expressCharge1 + otherCharge1;

                        } else {
                            for (int columnIndex = 19; columnIndex <= 36; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                        if (detailList != null && index < detailList.size()) {
                            ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorType() == null ? "" : detail.getErrorType().getLabel());
                            ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorCode() == null ? "" : detail.getErrorCode().getLabel());
                            ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getActionCode() == null ? "" : detail.getActionCode().getLabel());
                        } else {
                            for (int columnIndex = 39; columnIndex <= 41; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }


                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 9));

            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 13, 25));

            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalActualQty);

            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 27, 31));
            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalDetailCharge);
            ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTimelinessCharge);
            ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalUrgentCharge);
            ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCustomerPraiseFee);
            ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

            ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 37, 42));
        }
    }


    /**
     * 添加完工单Sheet 大于2000
     *
     * @param xBook
     * @param xStyle
     * @param orderMasterList
     */
    public void addCustomerChargeCompleteRptSheetMore2000(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCompletedOrderEntity> orderMasterList) {

        String xName = "完工单";
        Sheet xSheet = xBook.createSheet(xName);
        xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
        int rowIndex = 0;

        //====================================================绘制标题行============================================================
        Row titleRow = xSheet.createRow(rowIndex++);
        titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
        ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 41));

        //====================================================绘制表头============================================================
        //表头第一行
        Row headerFirstRow = xSheet.createRow(rowIndex++);
        headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 3));

        ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 4, 16));

        ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

        ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "到货时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

        ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "APP完成时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

        ExportExcel.createCell(headerFirstRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 20, 20));

        ExportExcel.createCell(headerFirstRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 21, 26));

        ExportExcel.createCell(headerFirstRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 27, 36));

        ExportExcel.createCell(headerFirstRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 37, 37));

        ExportExcel.createCell(headerFirstRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工结果");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 38, 38));
        ExportExcel.createCell(headerFirstRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 39, 41));
        ExportExcel.createCell(headerFirstRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结账日期");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 42, 42));

        //表头第二行
        Row headerSecondRow = xSheet.createRow(rowIndex++);
        headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
        ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
        ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");

        ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
        ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
        ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工条码");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

        ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
        ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");

        ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
        ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
        ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
        ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
        ExportExcel.createCell(headerSecondRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
        ExportExcel.createCell(headerSecondRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
        ExportExcel.createCell(headerSecondRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
        ExportExcel.createCell(headerSecondRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
        ExportExcel.createCell(headerSecondRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
        ExportExcel.createCell(headerSecondRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收合计");

        ExportExcel.createCell(headerSecondRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障类型");
        ExportExcel.createCell(headerSecondRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障现象");
        ExportExcel.createCell(headerSecondRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "故障处理");

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

        //====================================================绘制表格数据单元格============================================================
        int totalQty = 0;
        int totalActualQty = 0;
        double totalExpectCharge = 0.0;
        double totalCharge = 0.0;
        double totalDetailCharge = 0.0;
        double totalTimeLinessCharge = 0.0;
        double totalUrgentCharge = 0.0;
        double totalCustomerPraiseFee = 0.0;
        if (orderMasterList != null) {
            for (int i = 0; i < orderMasterList.size(); i++) {

                RPTCompletedOrderEntity orderMaster = orderMasterList.get(i);
                int rowSpan = orderMaster.getMaxRow() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();
                List<RPTOrderDetail> detailList = orderMaster.getDetails();

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());

                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem item = itemList.get(0);
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.joinStringList(item.getUnitBarCodes()));
                    totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                totalExpectCharge = totalExpectCharge + expectCharge;
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);

                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());

                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());


                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());

                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getArrivalDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getAppCompleteDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                RPTOrderDetail detail = null;
                if (detailList != null && detailList.size() > 0) {
                    detail = detailList.get(0);
                }

                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));
                int qty = (detail == null ? 0 : detail.getQty());
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                totalActualQty = totalActualQty + qty;

                double serviceCharge = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge);

                double materialCharge = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, materialCharge);

                double travelCharge = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, travelCharge);

                double expressCharge = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expressCharge);

                double otherCharge = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, otherCharge);
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge + materialCharge + travelCharge + expressCharge + otherCharge);
                totalDetailCharge = totalDetailCharge + serviceCharge + materialCharge + travelCharge + expressCharge + otherCharge;
                double customerTimeLinessCharge = orderMaster.getCustomerTimelinessCharge();
                totalTimeLinessCharge = totalTimeLinessCharge + customerTimeLinessCharge;
                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimeLinessCharge);

                double customerUrgentCharge = orderMaster.getCustomerUrgentCharge();
                totalUrgentCharge = totalUrgentCharge + customerUrgentCharge;
                ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerUrgentCharge);

                double customerPraiseFeeCharge = orderMaster.getCustomerPraiseFee();
                totalCustomerPraiseFee = totalCustomerPraiseFee + customerPraiseFeeCharge;
                ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerPraiseFeeCharge);
                ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTotalCharge());

                totalCharge = (totalCharge + orderMaster.getCustomerTotalCharge());
                ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "完成");
                ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getAppCompleteType() == null ? "" : orderMaster.getAppCompleteType().getLabel());
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorType() == null ? "" : detail.getErrorType().getLabel());
                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorCode() == null ? "" : detail.getErrorCode().getLabel());
                ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getActionCode() == null ? "" : detail.getActionCode().getLabel());
                ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getChargeDate(), "yyyy-MM-dd HH:mm:ss"));

                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem item = itemList.get(index);
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.joinStringList(item.getUnitBarCodes()));
                            totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        } else {
                            for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                        ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                        ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                        ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                        ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getArrivalDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getAppCompleteDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                        if (detailList != null && index < detailList.size()) {
                            detail = detailList.get(index);

                            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));
                            qty = (detail == null ? 0 : detail.getQty());
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                            totalActualQty = totalActualQty + qty;

                            double serviceCharge1 = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge1);

                            double materialCharge1 = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, materialCharge1);

                            double travelCharge1 = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, travelCharge1);

                            double expressCharge1 = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expressCharge1);

                            double otherCharge1 = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, otherCharge1);

                            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceCharge1 + materialCharge1 + travelCharge1 + expressCharge1 + otherCharge1);
                            totalDetailCharge = totalDetailCharge + serviceCharge1 + materialCharge1 + travelCharge1 + expressCharge1 + otherCharge1;
                            ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");

                            ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                            ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                            ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");

                            ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "完成");
                            ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getAppCompleteType() == null ? "" : orderMaster.getAppCompleteType().getLabel());
                            ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorType() == null ? "" : detail.getErrorType().getLabel());
                            ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getErrorCode() == null ? "" : detail.getErrorCode().getLabel());
                            ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : detail.getActionCode() == null ? "" : detail.getActionCode().getLabel());
                            ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getChargeDate(), "yyyy-MM-dd HH:mm:ss"));
                        } else {
                            for (int columnIndex = 19; columnIndex <= 36; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 2));
            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单数");
            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMasterList.size());
            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 5, 9));

            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 13, 25));

            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalActualQty);

            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 27, 31));
            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalDetailCharge);
            ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTimeLinessCharge);
            ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalUrgentCharge);
            ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCustomerPraiseFee);
            ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

            ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 37, 42));
        }
    }

}
