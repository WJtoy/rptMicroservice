package com.kkl.kklplus.provider.rpt.customer.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTUncompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.search.RPTUncompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.mapper.UncompletedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
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

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtUncompletedOrderRptService extends RptBaseService {

    @Resource
    private UncompletedOrderRptMapper uncompletedOrderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private AreaCacheService areaCacheService;


    public Page<RPTUncompletedOrderEntity> getUncompletedOrderRptData(RPTUncompletedOrderSearch search) {

        Page<RPTUncompletedOrderEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null && search.getEndDt() != null) {
            search.setEndDate(new Date(search.getEndDt()));
            List<String> quarters = getQuarters(search.getEndDate());
            if (!quarters.isEmpty()) {
                search.setQuarters(quarters);
                PageHelper.startPage(search.getPageNo(), search.getPageSize());
                returnPage = uncompletedOrderRptMapper.getUncompletedOrder(search);
                Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();

                List<RPTOrderItem> orderItemList = Lists.newArrayList();
                Map<String, RPTDict> pendingTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PENDING_TYPE);
                Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
                Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
                RPTCustomer customer = null;
                RPTArea area = null;
                String userName = null;
                List<RPTOrderItem> orderItems;
                Set<Long> customerIds = Sets.newHashSet();
                Set<Long> userIds = Sets.newHashSet();
                int dataSourceValue;
                for (RPTUncompletedOrderEntity item : returnPage) {
                    customerIds.add(item.getCustomer().getId());
                    userIds.add(item.getCreateBy().getId());
                    area = areaMap.get(item.getAreaId());
                    if (area != null) {
                        item.setUserAddress(area.getFullName() + " " + item.getUserAddress());
                    }
                    if(!pendingTypeMap.isEmpty()){
                        if (item.getPendingType() != null && item.getPendingType().getValue() != null) {
                            item.setPendingType(pendingTypeMap.get(item.getPendingType().getValue()));
                        }
                    }
                    if(!statusMap.isEmpty()){
                        if(item.getStatus() != null && item.getStatus().getValue() != null){
                            item.setStatus(statusMap.get(item.getStatus().getValue()));
                        }
                    }
                    dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
                    if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                        String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                        item.getShop().setLabel(shopName);
                    }
//                    orderItems = RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson());
                    orderItems = RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb());
                    item.setOrderItemPb(null);
                    item.setItems(orderItems);
                    orderItemList.addAll(item.getItems());
                }
                String[] fields = new String[]{"id", "name"};
                Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fields));
                Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
                for (RPTUncompletedOrderEntity item : returnPage) {
                    customer = customerMap.get(item.getCustomer().getId());
                    if (customer != null) {
                        if (StringUtils.isNotBlank(customer.getName())) {
                            item.getCustomer().setName(customer.getName());
                        }
                    }
                    userName = userNameMap.get(item.getCreateBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCreateBy().setName(userName);
                    }
                }

                RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
            }
        }
        return returnPage;
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
        RPTUncompletedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTUncompletedOrderSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getCustomerId() != null && searchCondition.getEndDate() != null) {
            List<String> quarters = getQuarters(searchCondition.getEndDate());
            if (!quarters.isEmpty()) {
                searchCondition.setQuarters(quarters);
                Integer rowCount = uncompletedOrderRptMapper.hasReportData(searchCondition);
                result = rowCount > 0;
            }
        }
        return result;
    }


    public SXSSFWorkbook UncompletedOrderExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTUncompletedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTUncompletedOrderSearch.class);
            Page<RPTUncompletedOrderEntity> list = getUncompletedOrderRptData(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(5000);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;


            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 20));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 2, 2));


            ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 3, 14));

            ExportExcel.createCell(headerFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 15, 15));

            ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户货款");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 16, 17));

            ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

            ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "停滞原因");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单金额");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "冻结金额");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            int totalQty = 0;
            double totalCharge = 0.0;
            double totalExpectCharge = 0.0;
            double totalBlockCharge = 0.0;

            if (list != null) {
                for (int i = 0; i < list.size(); i++) {

                    RPTUncompletedOrderEntity orderMaster = list.get(i);
                    int rowSpan = orderMaster.getMaxRow() - 1;
                    List<RPTOrderItem> itemList = orderMaster.getItems();

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                    }

                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                    }

                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                    }

                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    if (itemList != null && itemList.size() > 0) {
                        RPTOrderItem item = itemList.get(0);
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCharge());
                        totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        totalCharge = totalCharge + (item.getCharge() == null ? 0.0 : item.getCharge());
                    } else {
                        for (int columnIndex = 5; columnIndex <= 10; columnIndex++) {
                            ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                        }
                    }

                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                    }

                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                    }

                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                    }

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }

                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }

                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getExpectCharge());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }

                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getBlockedCharge());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }

                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus()==null ? "": orderMaster.getStatus().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getPendingType() == null ? "" : orderMaster.getPendingType().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                    }

                    if (rowSpan > 0) {
                        for (int index = 1; index <= rowSpan; index++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            if (itemList != null && index < itemList.size()) {
                                RPTOrderItem item = itemList.get(index);

                                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCharge());
                                totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                                totalCharge = totalCharge + (item.getCharge() == null ? 0.0 : item.getCharge());
                            } else {
                                for (int columnIndex = 5; columnIndex <= 10; columnIndex++) {
                                    ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                }
                            }
                        }
                    }
                }
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 8));

                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 11, 15));

                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBlockCharge);

                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 18, 19));
            }

        } catch (Exception e) {
            log.error("【UncompletedOrderRptService.UncompletedOrderExport】未完工单明细表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
