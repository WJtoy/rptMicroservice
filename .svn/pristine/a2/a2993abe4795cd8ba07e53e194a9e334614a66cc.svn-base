package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.b2bcenter.md.B2BDataSourceEnum;
import com.kkl.kklplus.entity.rpt.RPTCustomerNewOrderDailyRptEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.mapper.CustomerOrderDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
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

import javax.annotation.Resource;
import java.util.*;

/**
 * 客户每日下单明细
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerOrderMonthRptService extends RptBaseService {

    @Resource
    private CustomerOrderDailyRptMapper customerOrderDailyRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    /**
     * 获取客户每日下单列表
     */
    public Page<RPTCustomerNewOrderDailyRptEntity> getCustomerNewOrderMonthList(RPTSearchCondtion rptSearchCondtion) {

        Date date = DateUtils.longToDate(rptSearchCondtion.getBeginDate());
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        String quarter = DateUtils.getQuarter(endDate);

        if (rptSearchCondtion.getPageNo() != null && rptSearchCondtion.getPageSize() != null) {
            PageHelper.startPage(rptSearchCondtion.getPageNo(), rptSearchCondtion.getPageSize());
        }
        Page<RPTCustomerNewOrderDailyRptEntity> customerNewOrderDailyList = customerOrderDailyRptMapper.getCustomerNewOrderDailyList(rptSearchCondtion.getCustomerId(),
                rptSearchCondtion.getSalesId(), beginDate, endDate, quarter,rptSearchCondtion.getSubFlag());

        List<RPTOrderItem> orderItemList = Lists.newArrayList();
        Map<String, RPTDict> statusDictMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
        Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();
        Set<Long> customerIds = new HashSet<>();
        List<Long> createByIds = Lists.newArrayList();
        for (RPTCustomerNewOrderDailyRptEntity entity : customerNewOrderDailyList) {
            //设置下单信息
//            entity.setOrderItems(RPTOrderItemUtils.fromOrderItemsJson(entity.getOrderItemJson()));
            entity.setOrderItems(RPTOrderItemPbUtils.fromOrderItemsNewBytes(entity.getOrderItemPb()));
            orderItemList.addAll(entity.getOrderItems());
            entity.setOrderItemPb(null);
            entity.setOrderItemJson("");
            if (!customerIds.contains(entity.getCustomerId())) {
                customerIds.add(entity.getCustomerId());
            }
            if (!createByIds.contains(entity.getOrderCreateById())) {
                createByIds.add(entity.getOrderCreateById());
            }
            //设置订单状态
            if (statusDictMap != null && statusDictMap.get(String.valueOf(entity.getStatusValue())) != null && statusDictMap.get(String.valueOf(entity.getStatusValue())).getLabel() != null) {
                entity.setStatusLabel(statusDictMap.get(String.valueOf(entity.getStatusValue())).getLabel());
            }
            //设置店铺名称
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

        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
        Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(createByIds);
        for (RPTCustomerNewOrderDailyRptEntity entity : customerNewOrderDailyList) {
            //设置客户名称
            if (customerMap != null && customerMap.get(entity.getCustomerId()) != null) {
                entity.setCustomerName(customerMap.get(entity.getCustomerId()).getName());
            }
            //设置下单人
            if (namesByUserMap != null && namesByUserMap.get(entity.getOrderCreateById()) != null) {
                entity.setOrderCreateByName(namesByUserMap.get(entity.getOrderCreateById()));
            }
        }
        return customerNewOrderDailyList;
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
            Integer rowCount = customerOrderDailyRptMapper.hasReportData(searchCondition.getCustomerId(),
                    searchCondition.getSalesId(), beginDate, endDate, quarter,searchCondition.getSubFlag());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 创建Excel表格
     */
    public SXSSFWorkbook exportCustomerNewOrderMonthRpt(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTSearchCondtion searchCondtion = redisGsonService.fromJson(searchConditionJson, RPTSearchCondtion.class);
            List<RPTCustomerNewOrderDailyRptEntity> list = getCustomerNewOrderMonthList(searchCondtion);

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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 17));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 3));

            ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 4, 15));

            ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 16, 16));

            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));


            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            int totalCount = 0;
            double totalBlocked = 0d;
            int rowNumber = 0;
            for (RPTCustomerNewOrderDailyRptEntity orderMaster : list) {
                List<RPTOrderItem> orderItems = orderMaster.getOrderItems();
                rowNumber++;
                int rowSpan = orderMaster.getMaxRow() - 1;
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                }
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomerName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                }
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                }
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShopName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                }
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                }
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderCreateByName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                }
                if (orderItems != null && orderItems.size() > 0) {
                    RPTOrderItem item = orderItems.get(0);
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    totalCount = totalCount + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 6; columnIndex <= 10; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge());
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                }
                totalBlocked = totalBlocked + expectCharge;
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                }
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                }
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getServicePhone()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                }
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getFullServiceAddress()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                }
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (DateUtils.formatDate(orderMaster.getOrderCreateDate(), "yyyy-MM-dd HH:mm:ss")));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                }
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (orderMaster.getStatusLabel()));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                }


                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);

                        if (orderItems != null && index < orderItems.size()) {
                            RPTOrderItem item = orderItems.get(index);
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            totalCount = totalCount + StringUtils.toInteger(item.getQty());
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
            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);
            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBlocked);
            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 12, 17));
        } catch (Exception e) {
            log.error("客户每月下单明细报表写入excel失败:{}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
