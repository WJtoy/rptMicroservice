package com.kkl.kklplus.provider.rpt.customer.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.CustomerWriteOffRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
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
 * 客户退补单
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCustomerWriteOffRptService extends RptBaseService {

    @Resource
    private CustomerWriteOffRptMapper customerWriteOffRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;



    /**
     * 分页查询退补单明细
     */
    public Page<RPTCustomerWriteOffEntity> getCustomerWriteOffListByPaging(RPTCustomerWriteOffSearch search) {
        Page<RPTCustomerWriteOffEntity> returnPage = null;
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            search.setSystemId(RptCommonUtils.getSystemId());
            returnPage = customerWriteOffRptMapper.getCustomerWriteOffListByPaging(search);
            if (!returnPage.isEmpty()) {
                for (RPTCustomerWriteOffEntity item : returnPage) {
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
                    if(item.getDetails() != null && item.getDetails().size() > 0 ){
                        item.getDetails().get(0).setServiceCharge(item.getServiceCharge());
                        item.getDetails().get(0).setMaterialCharge(item.getMaterialCharge());
                        item.getDetails().get(0).setExpressCharge(item.getExpressCharge());
                        item.getDetails().get(0).setTravelCharge(item.getTravelCharge());
                        item.getDetails().get(0).setOtherCharge(item.getOtherCharge());
                    }else {
                        List<RPTOrderDetail> details = new ArrayList<>();
                        RPTOrderDetail detail = new RPTOrderDetail();
                        detail.setServiceCharge(item.getServiceCharge());
                        detail.setMaterialCharge(item.getMaterialCharge());
                        detail.setExpressCharge(item.getExpressCharge());
                        detail.setTravelCharge(item.getTravelCharge());
                        detail.setOtherCharge(item.getOtherCharge());
                        details.add(detail);
                        item.setDetails(details);
                    }
                    item.setOrderDetailPb(null);
                }
            }
        }
        return returnPage;
    }

    /**
     * 查询退补单明细
     */
    public List<RPTCustomerWriteOffEntity> getCustomerWriteOffList(RPTCustomerChargeSearch search) {
        List<RPTCustomerWriteOffEntity> returnList = null;
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
            returnList = customerWriteOffRptMapper.getCustomerWriteOffList(systemId, beginDateTime, endDateTime, search.getCustomerId(), quarter);
            if (!returnList.isEmpty()) {
                for (RPTCustomerWriteOffEntity item : returnList) {
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
                    if(item.getDetails() != null && item.getDetails().size() > 0 ){
                        item.getDetails().get(0).setServiceCharge(item.getServiceCharge());
                        item.getDetails().get(0).setMaterialCharge(item.getMaterialCharge());
                        item.getDetails().get(0).setExpressCharge(item.getExpressCharge());
                        item.getDetails().get(0).setTravelCharge(item.getTravelCharge());
                        item.getDetails().get(0).setOtherCharge(item.getOtherCharge());
                    }else {
                        List<RPTOrderDetail> details = new ArrayList<>();
                        RPTOrderDetail detail = new RPTOrderDetail();
                        detail.setServiceCharge(item.getServiceCharge());
                        detail.setMaterialCharge(item.getMaterialCharge());
                        detail.setExpressCharge(item.getExpressCharge());
                        detail.setTravelCharge(item.getTravelCharge());
                        detail.setOtherCharge(item.getOtherCharge());
                        details.add(detail);
                        item.setDetails(details);
                    }
                    item.setOrderDetailPb(null);
                }
            }
        }
        return returnList;
    }

    /**
     * 添加退补单Sheet
     *
     * @param xBook
     * @param xStyle
     * @param list
     */
    public void addCustomerWriteOffRptSheet(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCustomerWriteOffEntity> list) {

        String xName = "退补单";
        Sheet xSheet = xBook.createSheet(xName);
        xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
        int rowIndex = 0;

        //====================================================绘制标题行============================================================
        Row titleRow = xSheet.createRow(rowIndex++);
        titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
        ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 33));

        //====================================================绘制表头============================================================
        //表头第一行
        Row headerFirstRow = xSheet.createRow(rowIndex++);
        headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

        ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 2, 13));

        ExportExcel.createCell(headerFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 14, 14));

        ExportExcel.createCell(headerFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成日期");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 15, 15));

        ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 16, 21));

        ExportExcel.createCell(headerFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 22, 30));

        ExportExcel.createCell(headerFirstRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 31, 31));

        ExportExcel.createCell(headerFirstRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退补描述");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 32, 32));

        //表头第二行
        Row headerSecondRow = xSheet.createRow(rowIndex++);
        headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
        ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
        ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
        ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
        ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
        ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");

        ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
        ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
        ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
        ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
        ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
        ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效费");
        ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
        ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
        ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付合计");

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

        //====================================================绘制表格数据单元格============================================================
        int totalQty = 0;
        double totalExpectCharge = 0.0;
        double totalCharge = 0.0;
        double praiseFee = 0.0;
        double timelinessCharge = 0.0;
        double urgentCharge = 0.0;
        if (list != null) {
            for (int i = 0; i < list.size(); i++) {

                RPTCustomerWriteOffEntity orderMaster = list.get(i);
                int rowSpan = orderMaster.getMaxRow() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();
                List<RPTOrderDetail> detailList = orderMaster.getDetails();
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                }


                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                }

                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                }
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                }
                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem item = itemList.get(0);
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 4; columnIndex <= 8; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                totalExpectCharge = totalExpectCharge + orderMaster.getExpectCharge();
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getExpectCharge());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 9, 9));
                }

                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 10, 10));
                }

                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 11, 11));
                }

                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                }

                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                }

                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                }

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                }
                ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
                ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
                ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
                ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
                ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
                ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
                if (detailList != null && detailList.size() > 0) {
                    RPTOrderDetail item = detailList.get(0);
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceTimes());
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceType() == null ? "" : item.getServiceType().getName());
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProduct() == null ? "" : item.getProduct().getName());
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceCharge());
                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getMaterialCharge());
                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getExpressCharge());
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTravelCharge());
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getOtherCharge());
                } else {
                    for (int columnIndex = 16; columnIndex <= 26; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTimelinessCharge());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 27, 27));
                }
                timelinessCharge = timelinessCharge + orderMaster.getTimelinessCharge();
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUrgentCharge());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 28, 28));
                }
                urgentCharge = urgentCharge + orderMaster.getUrgentCharge();
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getPraiseFee());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 29, 29));
                }
                praiseFee = praiseFee + orderMaster.getPraiseFee();
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTotalCharge());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 30, 30));
                }
                totalCharge = totalCharge + orderMaster.getTotalCharge();

                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : "已结账");
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 31, 31));
                }
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());
                if (rowSpan > 0) {
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 32, 32));
                }

                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem item = itemList.get(index);

                            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        } else {
                            for (int columnIndex = 4; columnIndex <= 8; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        if (detailList != null && index < detailList.size()) {
                            RPTOrderDetail item = detailList.get(index);
                            ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceTimes());
                            ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceType() == null ? "" : item.getServiceType().getName());
                            ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProduct() == null ? "" : item.getProduct().getName());
                            ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceCharge());
                            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getMaterialCharge());
                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getExpressCharge());
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTravelCharge());
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getOtherCharge());
                        } else {
                            for (int columnIndex = 16; columnIndex <= 26; columnIndex++) {
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

            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 10, 26));

            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, timelinessCharge);
            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, urgentCharge);
            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);
            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 31, 32));
        }
    }


    /**
     * 添加退补单Sheet 大于2000
     *
     * @param xBook
     * @param xStyle
     * @param orderMasterList
     */
    public void addCustomerWriteOffRptSheetMore2000(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTCustomerWriteOffEntity> orderMasterList) {

        String xName = "退补单";
        Sheet xSheet = xBook.createSheet(xName);
        xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
        int rowIndex = 0;

        //====================================================绘制标题行============================================================
        Row titleRow = xSheet.createRow(rowIndex++);
        titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
        ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 33));

        //====================================================绘制表头============================================================
        //表头第一行
        Row headerFirstRow = xSheet.createRow(rowIndex++);
        headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

        ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 2, 13));

        ExportExcel.createCell(headerFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 14, 14));

        ExportExcel.createCell(headerFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成日期");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 15, 15));

        ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 16, 21));

        ExportExcel.createCell(headerFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 22, 30));

        ExportExcel.createCell(headerFirstRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 31, 31));

        ExportExcel.createCell(headerFirstRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退补描述");
        xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 32, 32));

        //表头第二行
        Row headerSecondRow = xSheet.createRow(rowIndex++);
        headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
        ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
        ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
        ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
        ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
        ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
        ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
        ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
        ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");

        ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
        ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
        ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
        ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
        ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
        ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效费");
        ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
        ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
        ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付合计");

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

        //====================================================绘制表格数据单元格============================================================
        int totalQty = 0;

        double totalExpectCharge = 0.0;
        double totalCharge = 0.0;
        double praiseFee = 0.0;
        double timelinessCharge = 0.0;
        double urgentCharge = 0.0;
        if (orderMasterList != null) {
            for (int i = 0; i < orderMasterList.size(); i++) {

                RPTCustomerWriteOffEntity orderMaster = orderMasterList.get(i);
                int rowSpan = orderMaster.getMaxRow() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();
                List<RPTOrderDetail> detailList = orderMaster.getDetails();
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);

                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());

                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());

                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());

                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem item = itemList.get(0);
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 4; columnIndex <= 8; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                totalExpectCharge = totalExpectCharge + orderMaster.getExpectCharge();
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getExpectCharge());
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());

                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());

                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());

                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                if (detailList != null && detailList.size() > 0) {
                    RPTOrderDetail item = detailList.get(0);
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceTimes());
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceType() == null ? "" : item.getServiceType().getName());
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProduct() == null ? "" : item.getProduct().getName());
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceCharge());
                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getMaterialCharge());
                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getExpressCharge());
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTravelCharge());
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getOtherCharge());
                } else {
                    for (int columnIndex = 16; columnIndex <= 26; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTimelinessCharge());
                timelinessCharge = timelinessCharge + orderMaster.getTimelinessCharge();

                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUrgentCharge());
                urgentCharge = urgentCharge + orderMaster.getUrgentCharge();

                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getPraiseFee());
                praiseFee = praiseFee + orderMaster.getPraiseFee();

                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTotalCharge());

                totalCharge = totalCharge + orderMaster.getTotalCharge();

                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : "已结账");


                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());


                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);

                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());

                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem item = itemList.get(index);

                            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        } else {
                            for (int columnIndex = 4; columnIndex <= 8; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                        ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());

                        ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());

                        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());

                        ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));

                        ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));

                        if (detailList != null && index < detailList.size()) {
                            RPTOrderDetail item = detailList.get(index);
                            ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceTimes());
                            ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceType() == null ? "" : item.getServiceType().getName());
                            ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProduct() == null ? "" : item.getProduct().getName());
                            ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getServiceCharge());
                            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getMaterialCharge());
                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getExpressCharge());
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTravelCharge());
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getOtherCharge());
                        } else {
                            for (int columnIndex = 16; columnIndex <= 26; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                        ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                        ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "0");
                    }
                }
            }
            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 0));
            ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "单数");
            ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMasterList.size());
            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 3, 7));

            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);

            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 10, 26));

            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, timelinessCharge);
            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, urgentCharge);
            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 31, 32));
        }
    }
}
