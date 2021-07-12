package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.b2bcenter.md.B2BDataSourceEnum;
import com.kkl.kklplus.entity.rpt.RPTCustomerFrozenDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.mapper.CustomerFrozenDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
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

import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerFrozenMonthRptService extends RptBaseService {

    @Autowired
    private CustomerFrozenDailyRptMapper customerFrozenDailyRptMapper;

    /**
     * 获取客户每日下单列表
     */
    public Page<RPTCustomerFrozenDailyEntity> getCustomerFrozenMonthList(RPTSearchCondtion rptSearchCondtion) {

        Date date = DateUtils.longToDate(rptSearchCondtion.getBeginDate());
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        String quarter = DateUtils.getQuarter(endDate);

        if (rptSearchCondtion.getPageNo() != null && rptSearchCondtion.getPageSize() != null) {
            PageHelper.startPage(rptSearchCondtion.getPageNo(), rptSearchCondtion.getPageSize());
        }
        Page<RPTCustomerFrozenDailyEntity> customerFrozenDailyList = customerFrozenDailyRptMapper.getCustomerFrozenDailyList(rptSearchCondtion.getCustomerId(),
                rptSearchCondtion.getSalesId(), beginDate, endDate, quarter,rptSearchCondtion.getSubFlag());

        List<RPTOrderItem> orderItemList = Lists.newArrayList();
        Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
        Set<String> missedOrderIds = Sets.newHashSet();
        Set<Long> createByIds = Sets.newHashSet();
        for (RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList) {

            if (StringUtils.isBlank(entity.getOrderNo())) {
                missedOrderIds.add(entity.getCurrencyNo());
            } else {
                createByIds.add(entity.getOrderCreateById());
            }
        }



        List<List<String>> missedOrderIdsList = Lists.partition(Lists.newArrayList(missedOrderIds), 20);
        List<RPTCustomerFrozenDailyEntity> orders;
        List<RPTCustomerFrozenDailyEntity> missedOrderList = Lists.newArrayList();
        for (List<String> currencyNos : missedOrderIdsList) {
            orders = customerFrozenDailyRptMapper.getCustomerFrozenDailyIdsFromWebDB(currencyNos);
            if (orders != null && !orders.isEmpty()) {
                for (RPTCustomerFrozenDailyEntity order : orders) {
                    createByIds.add(order.getOrderCreateById());
                }
                missedOrderList.addAll(orders);
            }
        }

        Map<String, RPTCustomerFrozenDailyEntity> missedOrderMap = missedOrderList.stream().collect(Collectors.toMap(RPTCustomerFrozenDailyEntity::getOrderNo, i -> i));

        RPTCustomerFrozenDailyEntity missedOrder;
        for(RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList){
            if (StringUtils.isBlank(entity.getOrderNo())) {
                missedOrder = missedOrderMap.get(entity.getCurrencyNo());
                if(missedOrder != null){
                    entity.setOrderNo(missedOrder.getOrderNo());
                    entity.setParentBizOrderId(missedOrder.getParentBizOrderId());
                    entity.setOrderCreateById(missedOrder.getOrderCreateById());
                    entity.setShopId(missedOrder.getShopId());
                    entity.setDataSourceId(missedOrder.getDataSourceId());
                    //                    entity.setOrderItemJson(missedOrder.getOrderItemJson());
                    entity.setOrderItemPb(missedOrder.getOrderItemPb());
                }
            }

            //设置下单信息
//            entity.setOrderItems(RPTOrderItemUtils.fromOrderItemsJson(entity.getOrderItemJson()));
            entity.setOrderItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(entity.getOrderItemPb()));
            entity.setOrderItemPb(null);
            entity.setOrderItemJson("");
            orderItemList.addAll(entity.getOrderItems());

            if (B2BDataSourceEnum.isB2BDataSource(entity.getDataSourceId())) {
                String shopName = shopMaps.getValue1().get(String.format("%d:%s", entity.getDataSourceId(), entity.getShopId()));
                if (shopName != null) {
                    entity.setShopName(shopName);
                }
            } else {
                String shopName = shopMaps.getValue0().get(entity.getShopId());
                if (shopName != null) {
                    entity.setShopName(shopName);
                }
            }


    }
        RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));

        Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(createByIds));
        for (RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList) {
            //设置下单人
            if (namesByUserMap != null && namesByUserMap.get(entity.getOrderCreateById()) != null) {
                entity.setOrderCreateByName(namesByUserMap.get(entity.getOrderCreateById()));
            }
        }


        return customerFrozenDailyList;

    }



    /**
     * 获取客户每日冻结明细列表
     */
    public List<RPTCustomerFrozenDailyEntity> getCustomerFrozenDailyByList(RPTSearchCondtion rptSearchCondtion) {

        Date date = DateUtils.longToDate(rptSearchCondtion.getBeginDate());
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        String quarter = DateUtils.getQuarter(endDate);

        List<RPTCustomerFrozenDailyEntity> customerFrozenDailyList = customerFrozenDailyRptMapper.getCustomerFrozenDailyByList(rptSearchCondtion.getCustomerId(),
                rptSearchCondtion.getSalesId(), beginDate, endDate, quarter,rptSearchCondtion.getSubFlag());

        List<RPTOrderItem> orderItemList = Lists.newArrayList();
        Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
        Set<String> missedOrderIds = Sets.newHashSet();
        Set<Long> createByIds = Sets.newHashSet();
        for (RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList) {

            if (StringUtils.isBlank(entity.getOrderNo())) {
                missedOrderIds.add(entity.getCurrencyNo());
            } else {
                createByIds.add(entity.getOrderCreateById());
            }
        }



        List<List<String>> missedOrderIdsList = Lists.partition(Lists.newArrayList(missedOrderIds), 20);
        List<RPTCustomerFrozenDailyEntity> orders;
        List<RPTCustomerFrozenDailyEntity> missedOrderList = Lists.newArrayList();
        for (List<String> currencyNos : missedOrderIdsList) {
            orders = customerFrozenDailyRptMapper.getCustomerFrozenDailyIdsFromWebDB(currencyNos);
            if (orders != null && !orders.isEmpty()) {
                for (RPTCustomerFrozenDailyEntity order : orders) {
                    createByIds.add(order.getOrderCreateById());
                }
                missedOrderList.addAll(orders);
            }
        }

        Map<String, RPTCustomerFrozenDailyEntity> missedOrderMap = missedOrderList.stream().collect(Collectors.toMap(RPTCustomerFrozenDailyEntity::getOrderNo, i -> i));

        RPTCustomerFrozenDailyEntity missedOrder;
        for(RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList){
            if (StringUtils.isBlank(entity.getOrderNo())) {
                missedOrder = missedOrderMap.get(entity.getCurrencyNo());
                if(missedOrder != null){
                    entity.setOrderNo(missedOrder.getOrderNo());
                    entity.setParentBizOrderId(missedOrder.getParentBizOrderId());
                    entity.setOrderCreateById(missedOrder.getOrderCreateById());
                    entity.setShopId(missedOrder.getShopId());
                    entity.setDataSourceId(missedOrder.getDataSourceId());
                //                    entity.setOrderItemJson(missedOrder.getOrderItemJson());
                    entity.setOrderItemPb(missedOrder.getOrderItemPb());
                }
            }

            //设置下单信息
//            entity.setOrderItems(RPTOrderItemUtils.fromOrderItemsJson(entity.getOrderItemJson()));
            entity.setOrderItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(entity.getOrderItemPb()));
            entity.setOrderItemPb(null);
            entity.setOrderItemJson("");
            orderItemList.addAll(entity.getOrderItems());


            if (B2BDataSourceEnum.isB2BDataSource(entity.getDataSourceId())) {
                String shopName = shopMaps.getValue1().get(String.format("%d:%s", entity.getDataSourceId(), entity.getShopId()));
                if (shopName != null) {
                    entity.setShopName(shopName);
                }
            } else {
                String shopName = shopMaps.getValue0().get(entity.getShopId());
                if (shopName != null) {
                    entity.setShopName(shopName);
                }
            }


        }
        RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));

        Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(createByIds));
        for (RPTCustomerFrozenDailyEntity entity : customerFrozenDailyList) {
            //设置下单人
            if (namesByUserMap != null && namesByUserMap.get(entity.getOrderCreateById()) != null) {
                entity.setOrderCreateByName(namesByUserMap.get(entity.getOrderCreateById()));
            }
        }


        return customerFrozenDailyList;

    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTSearchCondtion searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSearchCondtion.class);
        if (searchCondition != null && searchCondition.getBeginDate() != null && searchCondition.getBeginDate() > 0) {
            Date queryDate = DateUtils.longToDate(searchCondition.getBeginDate());
            Date beginDate = DateUtils.getStartDayOfMonth(queryDate);
            Date endDate = DateUtils.getLastDayOfMonth(queryDate);
            String quarter = DateUtils.getQuarter(queryDate);
            Integer rowCount = customerFrozenDailyRptMapper.hasReportData(searchCondition.getCustomerId(),
                    searchCondition.getSalesId(), beginDate, endDate, quarter,searchCondition.getSubFlag());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 创建Excel表格
     */
    public SXSSFWorkbook exportCustomerFrozenMonthRpt(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTSearchCondtion searchCondtion = redisGsonService.fromJson(searchConditionJson, RPTSearchCondtion.class);
            List<RPTCustomerFrozenDailyEntity> list = getCustomerFrozenDailyByList(searchCondtion);

            String xName = reportTitle;//"客户每日下单明细（ yyyy年MM月dd日）";

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 11));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺");
            ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
            ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
            ExportExcel.createCell(headerFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "冻结金额");
            ExportExcel.createCell(headerFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "冻结描述");
            ExportExcel.createCell(headerFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "解冻金额");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            double totalBlockAmount = 0d;
            double  unBlockAmount = 0d;
            int rowNumber = 0;
            for (RPTCustomerFrozenDailyEntity orderMaster : list) {
                List<RPTOrderItem> orderItems = orderMaster.getOrderItems();
                rowNumber++;
                int rowSpan = orderItems.size() - 1;
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                }

                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                }
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShopName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                }
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                }
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderCreateByName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                }
                if (orderItems != null && orderItems.size() > 0) {
                    RPTOrderItem item = orderItems.get(0);
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                } else {
                    for (int columnIndex = 5; columnIndex <= 7; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }



               if(orderMaster.getCurrencyType() == 10){
                   ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getBlockAmount()));
                   totalBlockAmount = totalBlockAmount + orderMaster.getBlockAmount();
                   if (rowSpan > 0) {
                       xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 8, 8));
                   }
               }else {
                   ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0);
                   if (rowSpan > 0) {
                       xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 8, 8));
                   }
               }

                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getRemarks());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 9, 9));
                }

                if(orderMaster.getCurrencyType() == 20){
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getBlockAmount()));
                    unBlockAmount = unBlockAmount + orderMaster.getBlockAmount();
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 10, 10));
                    }
                }else {
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 10, 10));
                    }
                }

                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);

                        if (orderItems != null && index < orderItems.size()) {
                            RPTOrderItem item = orderItems.get(index);
                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                        } else {
                            for (int columnIndex = 5; columnIndex <= 7; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 7));
            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBlockAmount);

            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, unBlockAmount);
        } catch (Exception e) {
            log.error("客户每月冻结明细报表写入excel失败:{}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }



}
