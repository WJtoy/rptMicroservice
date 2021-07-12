package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.entity.rpt.search.RPTCancelledOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.json.adapater.OrderItemAdapterForCancelledOrderRpt;
import com.kkl.kklplus.provider.rpt.mapper.CancelledOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
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
import java.util.stream.Collectors;

/**
 * 工单退单明细
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CancelledOrderRptService extends RptBaseService {

    @Resource
    private CancelledOrderRptMapper cancelledOrderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    /**
     * 从Web数据库提取工单退单明细数据
     */
    private List<RPTCancelledOrderEntity> getCancelledOrderListFromWebDB(Date date) {
        List<RPTCancelledOrderEntity> result = Lists.newArrayList();
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        List<RPTCancelledOrderEntity> list = cancelledOrderRptMapper.getCancelledOrderListFromWebDB(beginDate, endDate);
        if (list != null && !list.isEmpty()) {
            Map<String, RPTDict> dataSourceMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_DATA_SOURCE);
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
            Map<String, RPTDict> cancelResponsibleMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_CANCEL_RESPONSIBLE);
            Map<Long, RPTUser> keFuMap = MSUserUtils.getMapByUserType(RPTUser.USER_TYPE_KEFU);
            Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
            List<RPTOrderItem> allOrderItems = Lists.newArrayList();
            Set<Long> customerIds = Sets.newHashSet();
            Set<Long> userIds = Sets.newHashSet();
            RPTDict dataSourceDict;
            RPTDict paymentTypeDict;
            RPTDict statusDict;
            RPTDict cancelResponsibleDict;
            RPTUser kefu;
            for (RPTCancelledOrderEntity item : list) {
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
                cancelResponsibleDict = cancelResponsibleMap.get(item.getCancelResponsible().getValue());
                if (cancelResponsibleDict != null && StringUtils.isNotBlank(cancelResponsibleDict.getLabel())) {
                    item.getCancelResponsible().setLabel(cancelResponsibleDict.getLabel());
                }
                int dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
                if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                    String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                    item.getShop().setLabel(shopName);
                }
                if(item.getKefu() != null && item.getKefu().getId() != null){
                    kefu = keFuMap.get(item.getKefu().getId());
                    if(kefu != null){
                        item.getKefu().setName(kefu.getName());
                        item.getKefu().setPhone(kefu.getPhone());
                    }
                }
//                item.setItems(RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson()));
                item.setItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb()));
                allOrderItems.addAll(item.getItems());

                customerIds.add(item.getCustomer().getId());
                userIds.add(item.getCreateBy().getId());
                if (item.getCancelApplyBy() != null && item.getCancelApplyBy().getId() != null) {
                    userIds.add(item.getCancelApplyBy().getId());
                }
                if (item.getCancelApproveBy() != null && item.getCancelApproveBy().getId() != null) {
                    userIds.add(item.getCancelApproveBy().getId());
                }
            }
            RPTOrderItemUtils.setOrderItemProperties(allOrderItems, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
            String[] fieldsArray = new String[]{"id", "name","code","salesId"};
            Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
            Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                    .map(i -> i.getSales().getId()).collect(Collectors.toSet());
            userIds.addAll(salesIds);
            Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
            RPTCustomer customer;
            String userName;
            for (RPTCancelledOrderEntity item : list) {
                customer = customerMap.get(item.getCustomer().getId());
                if (customer != null) {
                    if (StringUtils.isNotBlank(customer.getName())) {
                        item.getCustomer().setName(customer.getName());
                    }
                    if (StringUtils.isNotBlank(customer.getCode())) {
                        item.getCustomer().setCode(customer.getCode());
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
                if (item.getCancelApplyBy() != null && item.getCancelApplyBy().getId() != null) {
                    userName = userNameMap.get(item.getCancelApplyBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCancelApplyBy().setName(userName);
                    }
                }
                if (item.getCancelApproveBy() != null && item.getCancelApproveBy().getId() != null) {
                    userName = userNameMap.get(item.getCancelApproveBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCancelApproveBy().setName(userName);
                    }
                }
            }
            result.addAll(list);
        }
        return result;
    }

    /**
     * 根据消息队列从Web数据库提取工单退单明细数据
     */
    private RPTCancelledOrderEntity getCancelledOrderListFromWebMQ(Long orderId) {
            RPTCancelledOrderEntity item = cancelledOrderRptMapper.getCancelledOrderListFromWebMQ(orderId);
            Map<String, RPTDict> dataSourceMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_DATA_SOURCE);
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
            Map<String, RPTDict> cancelResponsibleMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_CANCEL_RESPONSIBLE);
            Map<Long, RPTUser> keFuMap = MSUserUtils.getMapByUserType(RPTUser.USER_TYPE_KEFU);
            Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
            List<RPTOrderItem> allOrderItems = Lists.newArrayList();
            Set<Long> customerIds = Sets.newHashSet();
            Set<Long> userIds = Sets.newHashSet();
            RPTDict dataSourceDict;
            RPTDict paymentTypeDict;
            RPTDict statusDict;
            RPTDict cancelResponsibleDict;
            RPTUser kefu;
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
                cancelResponsibleDict = cancelResponsibleMap.get(item.getCancelResponsible().getValue());
                if (cancelResponsibleDict != null && StringUtils.isNotBlank(cancelResponsibleDict.getLabel())) {
                    item.getCancelResponsible().setLabel(cancelResponsibleDict.getLabel());
                }
                int dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
                if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                    String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                    item.getShop().setLabel(shopName);
                }
                if(item.getKefu() != null && item.getKefu().getId() != null){
                    kefu = keFuMap.get(item.getKefu().getId());
                    if(kefu != null){
                        item.getKefu().setName(kefu.getName());
                        item.getKefu().setPhone(kefu.getPhone());

                    }
                }
//                item.setItems(RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson()));
                item.setItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb()));
                allOrderItems.addAll(item.getItems());

                customerIds.add(item.getCustomer().getId());
                userIds.add(item.getCreateBy().getId());
                if (item.getCancelApplyBy() != null && item.getCancelApplyBy().getId() != null) {
                    userIds.add(item.getCancelApplyBy().getId());
                }
                if (item.getCancelApproveBy() != null && item.getCancelApproveBy().getId() != null) {
                    userIds.add(item.getCancelApproveBy().getId());
                }

            RPTOrderItemUtils.setOrderItemProperties(allOrderItems, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
            String[] fieldsArray = new String[]{"id", "name","code","salesId"};
            Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
            Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                    .map(i -> i.getSales().getId()).collect(Collectors.toSet());
            userIds.addAll(salesIds);
            Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
            RPTCustomer customer;
            String userName;
            customer = customerMap.get(item.getCustomer().getId());
                if (customer != null) {
                    if (StringUtils.isNotBlank(customer.getName())) {
                        item.getCustomer().setName(customer.getName());
                    }
                    if (StringUtils.isNotBlank(customer.getCode())) {
                        item.getCustomer().setCode(customer.getCode());
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
                if (item.getCancelApplyBy() != null && item.getCancelApplyBy().getId() != null) {
                    userName = userNameMap.get(item.getCancelApplyBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCancelApplyBy().setName(userName);
                    }
                }
                if (item.getCancelApproveBy() != null && item.getCancelApproveBy().getId() != null) {
                    userName = userNameMap.get(item.getCancelApproveBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCancelApproveBy().setName(userName);
                    }
                }


        return item;
    }


    public void save(MQRPTOrderProcessMessage.RPTOrderProcessMessage msg){
        if(msg.getOrderId()<=0){
            throw new RuntimeException("退单/取消单订单id不能为空");
        }
        RPTCancelledOrderEntity rptCancelledOrderEntity = getCancelledOrderListFromWebMQ(msg.getOrderId());
        if (rptCancelledOrderEntity !=null) {
            int systemId = RptCommonUtils.getSystemId();
            String quarter = DateUtils.getQuarter(rptCancelledOrderEntity.getCloseDate());
            insertCancelledOrder(rptCancelledOrderEntity,quarter,systemId);
        }

    }

    private void insertCancelledOrder(RPTCancelledOrderEntity entity, String quarter, int systemId) {
        try {
            entity.setSystemId(systemId);
            entity.setQuarter(quarter);
            entity.setCreateDt(entity.getCreateDate().getTime());
            entity.setCancelApplyDt(entity.getCancelApplyDate().getTime());
            entity.setCloseDt(entity.getCloseDate().getTime());
//            String orderItemJson = OrderItemAdapterForCancelledOrderRpt.toOrderItemsJson(entity.getItems());
//            entity.setOrderItemJson(orderItemJson);
            entity.setOrderItemPb(RPTOrderItemPbUtils.toOrderItemsBytes(entity.getItems()));
            cancelledOrderRptMapper.insert(entity);
        } catch (Exception e) {
            log.error("【CancelledOrderRptService.insertCancelledOrder】OrderId: {}, errorMsg: {}", entity.getOrderId(), Exceptions.getStackTraceAsString(e));
        }

    }

    private Map<Long, Long> getCancelledOrderIdMap(Integer systemId, Date date) {
        Date beginDate = DateUtils.getDateStart(date);
        Date endDate = DateUtils.getDateEnd(date);
        String quarter = QuarterUtils.getSeasonQuarter(date);
        List<LongTwoTuple> tuples = cancelledOrderRptMapper.getCancelledOrderIds(quarter, systemId, beginDate.getTime(), endDate.getTime());
        if (tuples != null && !tuples.isEmpty()) {
            return tuples.stream().collect(Collectors.toMap(TwoTuple::getBElement, TwoTuple::getAElement));
        } else {
            return Maps.newHashMap();
        }
    }

//    /**
//     * 将指定日期的退单或取消单保存到中间表
//     */
//    public void saveCancelledOrdersToRptDB(Date date) {
//        if (date != null) {
//            List<RPTCancelledOrderEntity> list = getCancelledOrderListFromWebDB(date);
//            if (!list.isEmpty()) {
//                String quarter = QuarterUtils.getSeasonQuarter(date);
//                int systemId = RptCommonUtils.getSystemId();
//                for (RPTCancelledOrderEntity item : list) {
//                    insertCancelledOrder(item, quarter, systemId);
//                }
//            }
//        }
//    }

    /**
     * 将工单系统中有的而中间表中没有的退单或取消单保存到中间表
     */
    public void saveMissedCancelledOrdersToRptDB(Date date) {
        if (date != null) {
            List<RPTCancelledOrderEntity> list = getCancelledOrderListFromWebDB(date);
            if (!list.isEmpty()) {
                String quarter = QuarterUtils.getSeasonQuarter(date);
                int systemId = RptCommonUtils.getSystemId();
                Map<Long, Long> cancelledOrderIdMap = getCancelledOrderIdMap(systemId, date);
                Long primaryKeyId;
                for (RPTCancelledOrderEntity item : list) {
                    primaryKeyId = cancelledOrderIdMap.get(item.getOrderId());
                    if (primaryKeyId == null || primaryKeyId == 0) {
                        insertCancelledOrder(item, quarter, systemId);
                    }
                }
            }
        }
    }

    /**
     * 删除中间表中指定日期的退单或取消单明细
     */
    public void deleteCancelledOrdersFromRptDB(Date date) {
        if (date != null) {
            String quarter = QuarterUtils.getSeasonQuarter(date);
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            cancelledOrderRptMapper.deleteCancelledOrders(quarter, systemId, beginDate.getTime(), endDate.getTime());
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
                            saveMissedCancelledOrdersToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            saveMissedCancelledOrdersToRptDB(beginDate);
                            break;
                        case UPDATE:
                            deleteCancelledOrdersFromRptDB(beginDate);
                            saveMissedCancelledOrdersToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteCancelledOrdersFromRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CancelledOrderRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    /**
     * 分页查询退单/取消单明细
     */
    public Page<RPTCancelledOrderEntity> getCancelledOrderListByPaging(RPTCancelledOrderSearch search) {
        Page<RPTCancelledOrderEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            search.setSystemId(RptCommonUtils.getSystemId());
            returnPage = cancelledOrderRptMapper.getCancelledOrderListByPaging(search);
            if (!returnPage.isEmpty()) {
                for (RPTCancelledOrderEntity item : returnPage) {
                    item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                    item.setOrderItemPb(null);
                    item.setCreateDate(new Date(item.getCreateDt()));
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setCancelApplyDate(new Date(item.getCancelApplyDt()));
                }
            }
        }
        return returnPage;
    }

    /**
     * 查询对账单中退单/取消单
     */
    public List<RPTCancelledOrderEntity> getCancelledOrder(RPTCustomerChargeSearch search) {
        List<RPTCancelledOrderEntity> returnList = null;
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
            returnList = cancelledOrderRptMapper.getCancelledOrderList(systemId, beginDateTime, endDateTime, search.getCustomerId(), quarter);
            if (!returnList.isEmpty()) {
                for (RPTCancelledOrderEntity item : returnList) {
                    item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                    item.setOrderItemPb(null);
                    item.setCreateDate(new Date(item.getCreateDt()));
                    item.setCloseDate(new Date(item.getCloseDt()));
                    item.setCancelApplyDate(new Date(item.getCancelApplyDt()));
                }
            }
        }
        return returnList;
    }

    /**
     *   对账单  ----添加退单/取消单Sheet

     * @return
     */
    public void addCustomerChargeReturnCancelRptSheet(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCancelledOrderEntity> list) {
        try {
            String xName = "退单取消单";
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, "退单/取消单");
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 19));

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

            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

            ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单/取消\n时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

            ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

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
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            int totalQty = 0;
            double totalExpectCharge = 0.0;

            if (list != null) {
                for (int i = 0; i < list.size(); i++) {

                    RPTCancelledOrderEntity orderMaster = list.get(i);
                    int rowSpan = orderMaster.getRowCount() - 1;
                    List<RPTOrderItem> itemList = orderMaster.getItems();

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
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                    }
                    if (itemList != null && itemList.size() > 0) {
                        RPTOrderItem orderItem = itemList.get(0);
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                        totalQty = totalQty + orderItem.getQty();
                    } else {
                        for (int columnIndex = 6; columnIndex <= 10; columnIndex++) {
                            ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                        }
                    }

                    double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                    totalExpectCharge = totalExpectCharge + expectCharge;
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                    }
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                    }
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                    }

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }

                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }

                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }

                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                    }

                    if (rowSpan > 0) {
                        for (int index = 1; index <= rowSpan; index++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            if (itemList != null && index < itemList.size()) {
                                RPTOrderItem orderItem = itemList.get(index);

                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                                totalQty = totalQty + orderItem.getQty();
                            } else {
                                for (int columnIndex = 6; columnIndex <= 10; columnIndex++) {
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
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 12, 19));
            }
        } catch (Exception e) {
            log.error("【CancelledOrderRptService.addCustomerChargeReturnCancelRptSheet】对账单-退单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));

        }
    }

    /**
     *  对账单----添加退单/取消单Sheet 大于2000
     *
     * @param xBook
     * @param xStyle
     */
    public void addCustomerChargeReturnCancelRptSheetMore2000(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCancelledOrderEntity> list) {

        String xName = "退单取消单";
        Sheet xSheet = xBook.createSheet(xName);
        xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
        int rowIndex = 0;

        //====================================================绘制标题行============================================================
        Row titleRow = xSheet.createRow(rowIndex++);
        titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
        ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, "退单/取消单");
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 19));

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

        ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

        ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单/取消\n时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

        ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

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
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

        //====================================================绘制表格数据单元格============================================================
        int totalQty = 0;
        double totalExpectCharge = 0.0;

        if (list != null) {
            for (int i = 0; i < list.size(); i++) {

                RPTCancelledOrderEntity orderMaster = list.get(i);
                int rowSpan = orderMaster.getRowCount() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem orderItem = itemList.get(0);
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());

                    totalQty = totalQty + orderItem.getQty();
                } else {
                    for (int columnIndex = 6; columnIndex <= 10; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                totalExpectCharge = totalExpectCharge + expectCharge;
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem orderItem = itemList.get(index);

                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                            totalQty = totalQty + orderItem.getQty();
                        } else {
                            for (int columnIndex = 6; columnIndex <= 10; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                        ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                        ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                        ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                        ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                        ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                        ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 2));
            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单数");
            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, list.size());
            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 5, 9));
            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 12, 19));
        }

    }
    //end 对账单 退单/取消单

    // begin  订单退单明细
    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCancelledOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCancelledOrderSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginCloseDate() != null && searchCondition.getEndCloseDate() != null) {
            Integer rowCount = cancelledOrderRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook cancelledOrderRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCancelledOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCancelledOrderSearch.class);
            Page<RPTCancelledOrderEntity> list = getCancelledOrderListByPaging(searchCondition);

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 22));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerRow = xSheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headerRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服名称");
            ExportExcel.createCell(headerRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服电话");
            ExportExcel.createCell(headerRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "跟进业务员");
            ExportExcel.createCell(headerRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约业务员");
            ExportExcel.createCell(headerRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单申请人");
            ExportExcel.createCell(headerRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
            ExportExcel.createCell(headerRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编号");
            ExportExcel.createCell(headerRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            ExportExcel.createCell(headerRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headerRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headerRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单类型");
            ExportExcel.createCell(headerRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
            ExportExcel.createCell(headerRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单时间");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            Row dataRow = null;
            int totalCount = 0;
            double totalCharge = 0d;
            if (list != null) {
                for (int i = 0; i < list.size(); i++) {
                    RPTCancelledOrderEntity orderMaster = list.get(i);
                    int rowSpan = orderMaster.getRowCount() - 1;
                    List<RPTOrderItem> itemList = orderMaster.getItems();

                    dataRow = xSheet.createRow(rowIndex++);
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

                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getKefu().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }

                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getKefu().getPhone());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }

                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getSales().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                    }

                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getSales().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 6, 6));
                    }
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelApplyBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 7, 7));
                    }

                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 8, 8));
                    }

                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 9, 9));
                    }

                    RPTOrderItem orderItem = null;
                    if (itemList != null && itemList.size() > 0) {
                        orderItem = itemList.get(0);
                    }
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));

                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));

                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getProductSpec()));

                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? 0 : orderItem.getQty()));
                    totalCount = totalCount + (orderItem == null ? 0 : StringUtils.toInteger(orderItem.getQty()));


                    double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                    totalCharge = totalCharge + expectCharge;
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }

                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }

                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }

                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                    }

                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelResponsible() == null ? "" : orderMaster.getCancelResponsible().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 20, 20));
                    }

                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelApplyComment());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 21, 21));
                    }

                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 22, 22));
                    }

                    if (rowSpan > 0) {
                        for (int itemIndex = 1; itemIndex < itemList.size(); itemIndex++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            orderItem = itemList.get(itemIndex);

                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));

                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));

                            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getProductSpec()));

                            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? 0 : orderItem.getQty()));
                            totalCount = totalCount + (orderItem == null ? 0 : StringUtils.toInteger(orderItem.getQty()));
                        }
                    }
                }

                dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 12));

                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);

                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 15, 22));
            }

        } catch (Exception e) {
            log.error("【CancelledOrderRptService.cancelledOrderRptExport】订单退单明细报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


}
