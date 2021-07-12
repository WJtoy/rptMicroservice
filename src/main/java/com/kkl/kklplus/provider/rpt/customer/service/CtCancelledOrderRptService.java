package com.kkl.kklplus.provider.rpt.customer.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCancelledOrderSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CancelledOrderRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
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
import java.util.stream.Collectors;

/**
 * 工单退单明细
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCancelledOrderRptService extends RptBaseService {

    @Resource
    private CancelledOrderRptMapper cancelledOrderRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;


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

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 2,2 ));

            ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 3, 15));

            ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 16, 16));

            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单/取消\n时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

            ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

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
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

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

                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                    }
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                    }
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    if (itemList != null && itemList.size() > 0) {
                        RPTOrderItem orderItem = itemList.get(0);
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                        totalQty = totalQty + orderItem.getQty();
                    } else {
                        for (int columnIndex = 5; columnIndex <= 9; columnIndex++) {
                            ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                        }
                    }

                    double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                    totalExpectCharge = totalExpectCharge + expectCharge;
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 10, 10));
                    }
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                    }
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                    }

                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                    }

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }

                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }

                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }

                    if (rowSpan > 0) {
                        for (int index = 1; index <= rowSpan; index++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            if (itemList != null && index < itemList.size()) {
                                RPTOrderItem orderItem = itemList.get(index);

                                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                                totalQty = totalQty + orderItem.getQty();
                            } else {
                                for (int columnIndex = 5; columnIndex <= 9; columnIndex++) {
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
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 11, 18));
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

        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

        ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 2, 2));

        ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 3, 15));

        ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 16, 16));

        ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单/取消\n时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 17, 17));

        ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 18, 18));

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
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

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
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem orderItem = itemList.get(0);
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());

                    totalQty = totalQty + orderItem.getQty();
                } else {
                    for (int columnIndex = 5; columnIndex <= 9; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                totalExpectCharge = totalExpectCharge + expectCharge;
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getParentBizOrderId()));
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem orderItem = itemList.get(index);

                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getProductSpec());
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getBrand());
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderItem.getQty());
                            totalQty = totalQty + orderItem.getQty();
                        } else {
                            for (int columnIndex = 5; columnIndex <= 9; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                        ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                        ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                        ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel()));
                        ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                        ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCancelApplyComment()));
                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 1));
            ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单数");
            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, list.size());
            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 4, 8));
            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 11, 18));
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 19));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerRow = xSheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headerRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "跟进业务员");
            ExportExcel.createCell(headerRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约业务员");
            ExportExcel.createCell(headerRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单申请人");
            ExportExcel.createCell(headerRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编号");
            ExportExcel.createCell(headerRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            ExportExcel.createCell(headerRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headerRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headerRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单类型");
            ExportExcel.createCell(headerRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单描述");
            ExportExcel.createCell(headerRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退单时间");

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

                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getSales().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }

                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getSales().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelApplyBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                    }
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 6, 6));
                    }

                    RPTOrderItem orderItem = null;
                    if (itemList != null && itemList.size() > 0) {
                        orderItem = itemList.get(0);
                    }
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));

                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));

                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? "" : orderItem.getProductSpec()));

                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderItem == null ? 0 : orderItem.getQty()));
                    totalCount = totalCount + (orderItem == null ? 0 : StringUtils.toInteger(orderItem.getQty()));


                    double expectCharge = (orderMaster.getExpectCharge() == null ? 0.0 : orderMaster.getExpectCharge());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                    totalCharge = totalCharge + expectCharge;
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                    }
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                    }
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                    }

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }

                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }

                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }

                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelResponsible() == null ? "" : orderMaster.getCancelResponsible().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }

                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCancelApplyComment());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }

                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                    }

                    if (rowSpan > 0) {
                        for (int itemIndex = 1; itemIndex < itemList.size(); itemIndex++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            orderItem = itemList.get(itemIndex);

                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getServiceType() == null ? "" : orderItem.getServiceType().getName()));

                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getProduct() == null ? "" : orderItem.getProduct().getName()));

                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? "" : orderItem.getProductSpec()));

                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (orderItem == null ? 0 : orderItem.getQty()));
                            totalCount = totalCount + (orderItem == null ? 0 : StringUtils.toInteger(orderItem.getQty()));
                        }
                    }
                }

                dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 9));

                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCount);

                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 12, 19));
            }

        } catch (Exception e) {
            log.error("【CancelledOrderRptService.cancelledOrderRptExport】订单退单明细报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


}
