package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderDetailsEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.search.RPTCompletedOrderDetailsSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.mapper.CompletedOrderDetailsRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTEngineerPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTServicePointPbUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CompletedOrderNewDetailsService extends RptBaseService {

    @Resource
    private CompletedOrderDetailsRptMapper completedOrderDetailsRptMapper;

    //获取报表显示数据
    public Page<RPTCompletedOrderDetailsEntity> getCompletedOrderNewDetailsRptList(RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> page = new Page<>();
        search.setSystemId(RptCommonUtils.getSystemId());
        int completedOrderSum = completedOrderDetailsRptMapper.getCompletedOrderSum(search);
        int customerWriteOffSum = completedOrderDetailsRptMapper.getCustomerWriteOffSum(search);
        int servicePointWriteOffSum = completedOrderDetailsRptMapper.getServicePointWriteOffSum(search);
        int pageSize = search.getPageSize();
        int pageNum = search.getPageNo();
        int total = completedOrderSum + customerWriteOffSum + servicePointWriteOffSum;
        int pageCount = (total + pageSize - 1)/ pageSize;
        page.setTotal(total);
        page.setPages(pageCount);
        int completedPageCount = (completedOrderSum + pageSize - 1) / pageSize;
        int completedRemainder = completedOrderSum % pageSize;

        int customerWriteOffRemainder = (completedOrderSum + customerWriteOffSum) % pageSize;

        int customerWriteOffCount =  (completedOrderSum + customerWriteOffSum + pageSize -1) / pageSize;
        int beginLimit;

        if (pageNum < completedPageCount) {
            beginLimit = (pageNum - 1) * pageSize;
            search.setBeginLimit(beginLimit);
            search.setEndLimit(pageSize);
            page.addAll(getCompletedOrderList(search));
        } else if (pageNum == completedPageCount) {
            if (completedRemainder == 0) {
                beginLimit = (pageNum - 1) * pageSize;
                search.setBeginLimit(beginLimit);
                search.setEndLimit(pageSize);
                page.addAll(getCompletedOrderList(search));
            } else {
                beginLimit = (pageNum - 1) * pageSize;
                search.setBeginLimit(beginLimit);
                search.setEndLimit(pageSize);
                page.addAll(getCompletedOrderList(search));
                search.setEndLimit(pageSize - completedRemainder);
                search.setBeginLimit(0);
                if (customerWriteOffSum > 0) {
                   page.addAll(getCustomerWriteOffList(search));
                }
                if (completedRemainder + customerWriteOffSum < pageSize) {
                    search.setEndLimit(pageSize - customerWriteOffRemainder);
                    search.setBeginLimit(0);
                    page.addAll(getServicePointWriteOffList(search));
                }
            }
        } else if (pageNum < customerWriteOffCount) {
            beginLimit = ((pageNum - completedPageCount - 1) * pageSize) + pageSize - completedRemainder;
            search.setBeginLimit(beginLimit);
            search.setEndLimit(pageSize);
            page.addAll(getCustomerWriteOffList(search));
        } else if (pageNum == customerWriteOffCount) {
            if (customerWriteOffRemainder == 0) {
                beginLimit = ((pageNum - completedPageCount - 1) * pageSize) + pageSize - completedRemainder;
                search.setBeginLimit(beginLimit);
                search.setEndLimit(pageSize);
                page.addAll(getCustomerWriteOffList(search));
            } else {
                beginLimit = ((pageNum - completedPageCount - 1) * pageSize) + pageSize - completedRemainder;
                search.setBeginLimit(beginLimit);
                search.setEndLimit(pageSize);
                page.addAll(getCustomerWriteOffList(search));
                search.setEndLimit(pageSize - customerWriteOffRemainder);
                search.setBeginLimit(0);
                page.addAll(getServicePointWriteOffList(search));
            }
        } else {
            beginLimit = ((pageNum - customerWriteOffCount - 1) * pageSize) + pageSize - customerWriteOffRemainder;
            search.setBeginLimit(beginLimit);
            search.setEndLimit(pageSize);
            page.addAll(getServicePointWriteOffList(search));
        }
        return page;

    }


    /**
     * 分页查询客户退补单明细
     */
    private Page<RPTCompletedOrderDetailsEntity> getCustomerWriteOffList(RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> returnPage;
        returnPage = completedOrderDetailsRptMapper.getCustomerWriteOffList(search);
        if (!returnPage.isEmpty()) {
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            RPTDict paymentTypeDict;
            double totalCharge;
            for (RPTCompletedOrderDetailsEntity item : returnPage) {

                paymentTypeDict = paymentTypeMap.get(item.getPaymentType().getValue());
                if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                    item.getPaymentType().setLabel(paymentTypeDict.getLabel());
                }
                if (item.getCustomer().getContractDt() != 0) {
                    item.getCustomer().setContractDate(new Date(item.getCustomer().getContractDt()));
                }
                item.getStatus().setLabel("客户退补");
                item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                item.setOrderItemPb(null);
                item.setDetails(RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb()));
                if (item.getDetails() != null && item.getDetails().size() > 0) {
                    item.getDetails().get(0).setServiceCharge(item.getServiceCharge());
                    item.getDetails().get(0).setMaterialCharge(item.getMaterialCharge());
                    item.getDetails().get(0).setExpressCharge(item.getExpressCharge());
                    item.getDetails().get(0).setTravelCharge(item.getTravelCharge());
                    item.getDetails().get(0).setOtherCharge(item.getOtherCharge());
                } else {
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
                totalCharge = item.getServiceCharge() +item.getMaterialCharge()+item.getExpressCharge() + item.getTravelCharge() + item.getOtherCharge();
                item.setCustomerSubtotalCharge(totalCharge);
            }
        }
        return returnPage;
    }

    /**
     * 分页查询完工单明细
     */
    private Page<RPTCompletedOrderDetailsEntity> getCompletedOrderList(RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> returnPage;

        returnPage = completedOrderDetailsRptMapper.getCompletedOrderList(search);
        if (!returnPage.isEmpty()) {
            List<RPTOrderDetail> orderDetails;
            List<RPTServicePoint> servicePoints;
            List<RPTEngineer> engineers;
            Map<Long, RPTServicePoint> servicePointMap;
            Map<Long, RPTEngineer> engineerMap;
            RPTServicePoint servicePoint;
            RPTEngineer engineer;
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            RPTDict paymentTypeDict;
            for (RPTCompletedOrderDetailsEntity item : returnPage) {
                if (item.getCustomer().getContractDt() != 0) {
                    item.getCustomer().setContractDate(new Date(item.getCustomer().getContractDt()));
                }
                paymentTypeDict = paymentTypeMap.get(item.getPaymentType().getValue());
                if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                    item.getPaymentType().setLabel(paymentTypeDict.getLabel());
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

        return returnPage;
    }

    /**
     * 分页查询网点退补明细
     */
    private Page<RPTCompletedOrderDetailsEntity> getServicePointWriteOffList(RPTCompletedOrderDetailsSearch search) {
        Page<RPTCompletedOrderDetailsEntity> returnPage;

        returnPage = completedOrderDetailsRptMapper.getServicePointWriteOffList(search);
        if (!returnPage.isEmpty()) {
            List<RPTOrderDetail> orderDetails;
            Map<Long, RPTServicePoint> servicePointMap;
            Map<Long, RPTEngineer> engineerMap;
            RPTServicePoint servicePoint;
            RPTEngineer engineer;
            List<RPTServicePoint> servicePoints = new ArrayList<>();
            List<RPTEngineer> engineers = new ArrayList<>();
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
            RPTDict paymentTypeDict;
            double totalCharge;
            for (RPTCompletedOrderDetailsEntity item : returnPage) {
                paymentTypeDict = paymentTypeMap.get(item.getPaymentType().getValue());
                if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                    item.getPaymentType().setLabel(paymentTypeDict.getLabel());
                }
                item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                item.setOrderItemPb(null);
                item.getStatus().setLabel("网点退补");
                orderDetails = RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
                item.setDetails(orderDetails);
                item.setOrderDetailPb(null);

                servicePoints.add(RPTServicePointPbUtils.fromServicePointBytes(item.getServicePointPb()));
                if (!servicePoints.isEmpty()) {
                    item.setServicePoints(servicePoints);
                }
                engineers.add(RPTEngineerPbUtils.fromEngineerBytes(item.getEngineerPb()));
                if (!engineers.isEmpty()) {
                    item.setEngineers(engineers);
                }
                item.setServicePointPb(null);
                item.setEngineerPb(null);
                servicePointMap = servicePoints.stream().collect(Collectors.toMap(RPTBase::getId, i -> i,(key1, key2) -> key2));
                engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i,(key1, key2) -> key2));
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
                if (item.getDetails() != null && item.getDetails().size() > 0) {
                    item.getDetails().get(0).setEngineerServiceCharge(item.getServiceCharge());
                    item.getDetails().get(0).setEngineerMaterialCharge(item.getMaterialCharge());
                    item.getDetails().get(0).setEngineerExpressCharge(item.getExpressCharge());
                    item.getDetails().get(0).setEngineerTravelCharge(item.getTravelCharge());
                    item.getDetails().get(0).setEngineerOtherCharge(item.getOtherCharge());
                } else {
                    List<RPTOrderDetail> details = new ArrayList<>();
                    RPTOrderDetail detail = new RPTOrderDetail();
                    detail.setEngineerServiceCharge(item.getServiceCharge());
                    detail.setEngineerMaterialCharge(item.getMaterialCharge());
                    detail.setEngineerExpressCharge(item.getExpressCharge());
                    detail.setEngineerTravelCharge(item.getTravelCharge());
                    detail.setEngineerOtherCharge(item.getOtherCharge());
                    details.add(detail);
                    item.setDetails(details);
                }
                totalCharge = item.getServiceCharge() +item.getMaterialCharge()+item.getExpressCharge() + item.getTravelCharge() + item.getOtherCharge();
                item.setEngineerSubtotalCharge(totalCharge);

            }
        }

        return returnPage;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount;
            Integer completedOrderSum = completedOrderDetailsRptMapper.getCompletedOrderSum(searchCondition);
            Integer customerWriteOffSum = completedOrderDetailsRptMapper.getCustomerWriteOffSum(searchCondition);
            Integer servicePointWriteOffSum = completedOrderDetailsRptMapper.getServicePointWriteOffSum(searchCondition);
            rowCount = completedOrderSum + customerWriteOffSum + servicePointWriteOffSum;
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook completedOrderDetailsNewExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTCompletedOrderDetailsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCompletedOrderDetailsSearch.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 42));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);


            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 0, 1));

            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 2, 5));

            ExportExcel.createCell(headerFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "安维人员信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 6, 9));

            ExportExcel.createCell(headerFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 10, 16));

            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付安维费用");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 17, 31));

            ExportExcel.createCell(headerFirstRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 32, 41));

            ExportExcel.createCell(headerFirstRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退补描述");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 42, 42));

            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "姓名");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "电话");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
            ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
            ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
            ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
            ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
            ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效奖励");
            ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
            ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
            ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "互助基金");
            ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");
            ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台费");
            ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保金额");
            ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headerSecondRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付合计");
            ExportExcel.createCell(headerSecondRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
            ExportExcel.createCell(headerSecondRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
            ExportExcel.createCell(headerSecondRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
            ExportExcel.createCell(headerSecondRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
            ExportExcel.createCell(headerSecondRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
            ExportExcel.createCell(headerSecondRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
            ExportExcel.createCell(headerSecondRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headerSecondRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
            ExportExcel.createCell(headerSecondRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headerSecondRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收合计");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================

            int totalCount = 0;
            int totalActualCount = 0;
            double totalInCharge = 0d;
            double totalOutCharge = 0d;
            double engineerServiceChargeTotal = 0d;
            double engineerMaterialChargeTotal = 0d;
            double engineerTravelChargeTotal = 0d;
            double engineerDismantleChargeTotal = 0d;
            double engineerOtherChargeTotal = 0d;
            double engineerInsuranceChargeTotal = 0d;
            double engineerTimelinessChargeTotal = 0d;
            double engineerCustomerTimelinessChargeTotal = 0d;
            double engineerUrgentChargeTotal = 0d;
            double engineerTaxFeeTotal = 0d;
            double engineerInfoFeeTotal = 0d;
            double engineerDepositTotal = 0d;
            double engineerPraiseFeeTotal = 0d;
            double customerServiceChargeTotal = 0d;
            double customerPraiseFeeTotal = 0d;
            double customerMaterialChargeTotal = 0d;
            double customerTravelChargeTotal = 0d;
            double customerDismantleChargeTotal = 0d;
            double customerOtherChargeTotal = 0d;
            double customerTimelinessChargeTotal = 0d;
            double customerUrgentChargeTotal = 0d;
            double customerChargeSumTotal = 0d;
            double engineerChargeTotal = 0d;

            int listSize = 0;
            int endLimit = 5000;
            int beginLimit = 1;
            int rowNumber = 0;
            int finishQty = 0;
            int serviceWrite  = 0;
            int customerWrite = 0;

            do {
                searchCondition.setBeginLimit((beginLimit - 1) * endLimit);
                searchCondition.setEndLimit(endLimit);
                Page<RPTCompletedOrderDetailsEntity> list = getCompletedOrderList(searchCondition);
                if (list != null) {
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalActualCount, totalInCharge, totalOutCharge, engineerServiceChargeTotal, engineerMaterialChargeTotal, engineerTravelChargeTotal,
                                                                                         engineerDismantleChargeTotal, engineerOtherChargeTotal, engineerInsuranceChargeTotal, engineerTimelinessChargeTotal, engineerCustomerTimelinessChargeTotal,
                                                                                         engineerUrgentChargeTotal, engineerTaxFeeTotal, engineerInfoFeeTotal, engineerDepositTotal, engineerPraiseFeeTotal, customerServiceChargeTotal, customerPraiseFeeTotal, engineerChargeTotal,customerMaterialChargeTotal, customerTravelChargeTotal, customerDismantleChargeTotal , customerOtherChargeTotal, customerTimelinessChargeTotal, customerUrgentChargeTotal, customerChargeSumTotal,rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount =  dataExport.getTotalCount();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge =  dataExport.getTotalInCharge();
                    totalOutCharge =  dataExport.getTotalOutCharge();
                    engineerServiceChargeTotal = dataExport.getEngineerServiceChargeTotal();
                    engineerMaterialChargeTotal = dataExport.getEngineerMaterialChargeTotal();
                    engineerTravelChargeTotal = dataExport.getEngineerTravelChargeTotal();
                    engineerDismantleChargeTotal = dataExport.getEngineerDismantleChargeTotal();
                    engineerOtherChargeTotal = dataExport.getEngineerOtherChargeTotal();
                    engineerInsuranceChargeTotal = dataExport.getEngineerInsuranceChargeTotal();
                    engineerTimelinessChargeTotal = dataExport.getEngineerTimelinessChargeTotal();
                    engineerCustomerTimelinessChargeTotal = dataExport.getEngineerCustomerTimelinessChargeTotal();
                    engineerUrgentChargeTotal = dataExport.getEngineerUrgentChargeTotal();
                    engineerTaxFeeTotal = dataExport.getEngineerTaxFeeTotal();
                    engineerInfoFeeTotal = dataExport.getEngineerInfoFeeTotal();
                    engineerDepositTotal = dataExport.getEngineerDepositTotal();
                    engineerPraiseFeeTotal = dataExport.getEngineerPraiseFeeTotal();
                    customerServiceChargeTotal = dataExport.getCustomerServiceChargeTotal();
                    customerPraiseFeeTotal = dataExport.getCustomerPraiseFeeTotal();
                    engineerChargeTotal = dataExport.getEngineerChargeTotal();
                    customerMaterialChargeTotal = dataExport.getCustomerMaterialChargeTotal();
                    customerTravelChargeTotal = dataExport.getCustomerTravelChargeTotal();
                    customerDismantleChargeTotal = dataExport.getCustomerDismantleChargeTotal();
                    customerOtherChargeTotal = dataExport.getCustomerOtherChargeTotal();
                    customerTimelinessChargeTotal = dataExport.getCustomerTimelinessChargeTotal();
                    customerUrgentChargeTotal = dataExport.getCustomerUrgentChargeTotal();
                    customerChargeSumTotal = dataExport.getCustomerChargeSumTotal();
                    rowNumber = dataExport.getRowNumber();
                    listSize = list.size();
                }

                finishQty = finishQty + listSize;
                beginLimit++;
            } while (listSize == endLimit);
            beginLimit = 1;
            do {
                searchCondition.setBeginLimit((beginLimit - 1) * endLimit);
                searchCondition.setEndLimit(endLimit);
                Page<RPTCompletedOrderDetailsEntity> list = getCustomerWriteOffList(searchCondition);

                if (list != null) {
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalActualCount, totalInCharge, totalOutCharge, engineerServiceChargeTotal, engineerMaterialChargeTotal, engineerTravelChargeTotal,
                                                                                         engineerDismantleChargeTotal, engineerOtherChargeTotal, engineerInsuranceChargeTotal, engineerTimelinessChargeTotal, engineerCustomerTimelinessChargeTotal,
                                                                                         engineerUrgentChargeTotal, engineerTaxFeeTotal, engineerInfoFeeTotal, engineerDepositTotal, engineerPraiseFeeTotal, customerServiceChargeTotal, customerPraiseFeeTotal,engineerChargeTotal,customerMaterialChargeTotal, customerTravelChargeTotal, customerDismantleChargeTotal , customerOtherChargeTotal, customerTimelinessChargeTotal, customerUrgentChargeTotal, customerChargeSumTotal,rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount =  dataExport.getTotalCount();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge =  dataExport.getTotalInCharge();
                    totalOutCharge =  dataExport.getTotalOutCharge();
                    engineerServiceChargeTotal = dataExport.getEngineerServiceChargeTotal();
                    engineerMaterialChargeTotal = dataExport.getEngineerMaterialChargeTotal();
                    engineerTravelChargeTotal = dataExport.getEngineerTravelChargeTotal();
                    engineerDismantleChargeTotal = dataExport.getEngineerDismantleChargeTotal();
                    engineerOtherChargeTotal = dataExport.getEngineerOtherChargeTotal();
                    engineerInsuranceChargeTotal = dataExport.getEngineerInsuranceChargeTotal();
                    engineerTimelinessChargeTotal = dataExport.getEngineerTimelinessChargeTotal();
                    engineerCustomerTimelinessChargeTotal = dataExport.getEngineerCustomerTimelinessChargeTotal();
                    engineerUrgentChargeTotal = dataExport.getEngineerUrgentChargeTotal();
                    engineerTaxFeeTotal = dataExport.getEngineerTaxFeeTotal();
                    engineerInfoFeeTotal = dataExport.getEngineerInfoFeeTotal();
                    engineerDepositTotal = dataExport.getEngineerDepositTotal();
                    engineerPraiseFeeTotal = dataExport.getEngineerPraiseFeeTotal();
                    customerServiceChargeTotal = dataExport.getCustomerServiceChargeTotal();
                    customerPraiseFeeTotal = dataExport.getCustomerPraiseFeeTotal();
                    engineerChargeTotal = dataExport.getEngineerChargeTotal();
                    customerMaterialChargeTotal = dataExport.getCustomerMaterialChargeTotal();
                    customerTravelChargeTotal = dataExport.getCustomerTravelChargeTotal();
                    customerDismantleChargeTotal = dataExport.getCustomerDismantleChargeTotal();
                    customerOtherChargeTotal = dataExport.getCustomerOtherChargeTotal();
                    customerTimelinessChargeTotal = dataExport.getCustomerTimelinessChargeTotal();
                    customerUrgentChargeTotal = dataExport.getCustomerUrgentChargeTotal();
                    customerChargeSumTotal = dataExport.getCustomerChargeSumTotal();
                    rowNumber = dataExport.getRowNumber();
                    listSize =  list.size();
                }
                customerWrite = customerWrite + listSize;
                beginLimit++;
            } while (listSize == endLimit);
            beginLimit = 1;
            do {
                searchCondition.setBeginLimit((beginLimit - 1) * endLimit);
                searchCondition.setEndLimit(endLimit);
                Page<RPTCompletedOrderDetailsEntity> list = getServicePointWriteOffList(searchCondition);
                if (list != null) {
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalActualCount, totalInCharge, totalOutCharge, engineerServiceChargeTotal, engineerMaterialChargeTotal, engineerTravelChargeTotal,
                                                                                         engineerDismantleChargeTotal, engineerOtherChargeTotal, engineerInsuranceChargeTotal, engineerTimelinessChargeTotal, engineerCustomerTimelinessChargeTotal,
                                                                                         engineerUrgentChargeTotal, engineerTaxFeeTotal, engineerInfoFeeTotal, engineerDepositTotal, engineerPraiseFeeTotal, customerServiceChargeTotal, customerPraiseFeeTotal,engineerChargeTotal,customerMaterialChargeTotal, customerTravelChargeTotal, customerDismantleChargeTotal , customerOtherChargeTotal, customerTimelinessChargeTotal, customerUrgentChargeTotal, customerChargeSumTotal,rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount = dataExport.getTotalCount();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge = dataExport.getTotalInCharge();
                    totalOutCharge = dataExport.getTotalOutCharge();
                    engineerServiceChargeTotal = dataExport.getEngineerServiceChargeTotal();
                    engineerMaterialChargeTotal = dataExport.getEngineerMaterialChargeTotal();
                    engineerTravelChargeTotal = dataExport.getEngineerTravelChargeTotal();
                    engineerDismantleChargeTotal = dataExport.getEngineerDismantleChargeTotal();
                    engineerOtherChargeTotal = dataExport.getEngineerOtherChargeTotal();
                    engineerInsuranceChargeTotal = dataExport.getEngineerInsuranceChargeTotal();
                    engineerTimelinessChargeTotal = dataExport.getEngineerTimelinessChargeTotal();
                    engineerCustomerTimelinessChargeTotal = dataExport.getEngineerCustomerTimelinessChargeTotal();
                    engineerUrgentChargeTotal = dataExport.getEngineerUrgentChargeTotal();
                    engineerTaxFeeTotal = dataExport.getEngineerTaxFeeTotal();
                    engineerInfoFeeTotal = dataExport.getEngineerInfoFeeTotal();
                    engineerDepositTotal = dataExport.getEngineerDepositTotal();
                    engineerPraiseFeeTotal = dataExport.getEngineerPraiseFeeTotal();
                    customerServiceChargeTotal = dataExport.getCustomerServiceChargeTotal();
                    customerPraiseFeeTotal = dataExport.getCustomerPraiseFeeTotal();
                    engineerChargeTotal = dataExport.getEngineerChargeTotal();
                    customerMaterialChargeTotal = dataExport.getCustomerMaterialChargeTotal();
                    customerTravelChargeTotal = dataExport.getCustomerTravelChargeTotal();
                    customerDismantleChargeTotal = dataExport.getCustomerDismantleChargeTotal();
                    customerOtherChargeTotal = dataExport.getCustomerOtherChargeTotal();
                    customerTimelinessChargeTotal = dataExport.getCustomerTimelinessChargeTotal();
                    customerUrgentChargeTotal = dataExport.getCustomerUrgentChargeTotal();
                    customerChargeSumTotal = dataExport.getCustomerChargeSumTotal();
                    rowNumber = dataExport.getRowNumber();
                    listSize = list.size();
                }
                serviceWrite = serviceWrite + listSize;
                beginLimit++;
            } while (listSize == endLimit);

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 4));
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单数量");
                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, finishQty);
                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户退补单量");
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerWrite);
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点退补单量");
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceWrite);
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalActualCount);

                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerServiceChargeTotal);
                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerMaterialChargeTotal);
                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTravelChargeTotal);
                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDismantleChargeTotal);
                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerOtherChargeTotal);
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTimelinessChargeTotal);
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerCustomerTimelinessChargeTotal);
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerUrgentChargeTotal);
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerChargeTotal);
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInsuranceChargeTotal);
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTaxFeeTotal);
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInfoFeeTotal);
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDepositTotal);
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFeeTotal);
                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOutCharge);
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerServiceChargeTotal);
                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerMaterialChargeTotal);
                ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTravelChargeTotal);
                ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerDismantleChargeTotal);
                ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerOtherChargeTotal);
                ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTimelinessChargeTotal);
                ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerUrgentChargeTotal);
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerChargeSumTotal);

                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerPraiseFeeTotal);
                ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalInCharge);

                ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");


        } catch (Exception e) {
            log.error("【CompletedOrderDetailsService.completedOrderDetailsExport】完工单明细报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    private class writeCompletionDataExport {
        private Sheet xSheet;
        private Map<String, CellStyle> xStyle;
        private int rowIndex;
        private int totalCount;
        private int totalActualCount;
        private double totalInCharge;
        private double totalOutCharge;
        private double engineerServiceChargeTotal;
        private double engineerMaterialChargeTotal;
        private double engineerTravelChargeTotal;
        private double engineerDismantleChargeTotal;
        private double engineerOtherChargeTotal;
        private double engineerInsuranceChargeTotal;
        private double engineerTimelinessChargeTotal;
        private double engineerCustomerTimelinessChargeTotal;
        private double engineerUrgentChargeTotal;
        private double engineerTaxFeeTotal;
        private double engineerInfoFeeTotal;
        private double engineerDepositTotal;
        private double engineerPraiseFeeTotal;
        private double customerServiceChargeTotal;
        private double customerPraiseFeeTotal;
        private double engineerChargeTotal;
        private double customerMaterialChargeTotal;
        private double customerTravelChargeTotal;
        private double customerDismantleChargeTotal;
        private double customerOtherChargeTotal;
        private double customerTimelinessChargeTotal;
        private double customerUrgentChargeTotal;
        private double customerChargeSumTotal;
        private int rowNumber;
        private Page<RPTCompletedOrderDetailsEntity> list;

        private writeCompletionDataExport(Sheet xSheet, Map<String, CellStyle> xStyle, int rowIndex, int totalCount, int totalActualCount, double totalInCharge, double totalOutCharge, double engineerServiceChargeTotal,
                double engineerMaterialChargeTotal, double engineerTravelChargeTotal, double engineerDismantleChargeTotal, double engineerOtherChargeTotal, double engineerInsuranceChargeTotal, double engineerTimelinessChargeTotal, double engineerCustomerTimelinessChargeTotal, double engineerUrgentChargeTotal,
                double engineerTaxFeeTotal, double engineerInfoFeeTotal, double engineerDepositTotal, double engineerPraiseFeeTotal, double customerServiceChargeTotal, double customerPraiseFeeTotal,double engineerChargeTotal, double customerMaterialChargeTotal,double customerTravelChargeTotal, double customerDismantleChargeTotal,
                double customerOtherChargeTotal, double customerTimelinessChargeTotal, double customerUrgentChargeTotal, double customerChargeSumTotal,int rowNumber, Page<RPTCompletedOrderDetailsEntity> list) {

            this.xSheet = xSheet;
            this.xStyle = xStyle;
            this.rowIndex = rowIndex;
            this.totalCount = totalCount;
            this.totalActualCount = totalActualCount;
            this.totalInCharge = totalInCharge;
            this.totalOutCharge = totalOutCharge;
            this.engineerServiceChargeTotal = engineerServiceChargeTotal;
            this.engineerMaterialChargeTotal = engineerMaterialChargeTotal;
            this.engineerTravelChargeTotal = engineerTravelChargeTotal;
            this.engineerDismantleChargeTotal = engineerDismantleChargeTotal;
            this.engineerOtherChargeTotal = engineerOtherChargeTotal;
            this.engineerInsuranceChargeTotal = engineerInsuranceChargeTotal;
            this.engineerTimelinessChargeTotal = engineerTimelinessChargeTotal;
            this.engineerCustomerTimelinessChargeTotal = engineerCustomerTimelinessChargeTotal;
            this.engineerUrgentChargeTotal = engineerUrgentChargeTotal;
            this.engineerTaxFeeTotal = engineerTaxFeeTotal;
            this.engineerInfoFeeTotal = engineerInfoFeeTotal;
            this.engineerDepositTotal = engineerDepositTotal;
            this.engineerPraiseFeeTotal = engineerPraiseFeeTotal;
            this.customerServiceChargeTotal = customerServiceChargeTotal;
            this.customerPraiseFeeTotal = customerPraiseFeeTotal;
            this.engineerChargeTotal = engineerChargeTotal;
            this.customerMaterialChargeTotal = customerMaterialChargeTotal;
            this.customerTravelChargeTotal = customerTravelChargeTotal;
            this.customerDismantleChargeTotal = customerDismantleChargeTotal;
            this.customerOtherChargeTotal = customerOtherChargeTotal;
            this.customerTimelinessChargeTotal = customerTimelinessChargeTotal;
            this.customerUrgentChargeTotal = customerUrgentChargeTotal;
            this.customerChargeSumTotal = customerChargeSumTotal;
            this.rowNumber = rowNumber;
            this.list = list;
        }

        private int getRowIndex() {
            return rowIndex;
        }

        private int getTotalCount() {
            return totalCount;
        }

        private int getTotalActualCount() {
            return totalActualCount;
        }

        private double getEngineerChargeTotal() {
            return engineerChargeTotal;
        }

        private double getEngineerServiceChargeTotal() {
            return engineerServiceChargeTotal;
        }

        private double getEngineerMaterialChargeTotal() {
            return engineerMaterialChargeTotal;
        }

        private double getEngineerTravelChargeTotal() {
            return engineerTravelChargeTotal;
        }

        private double getEngineerDismantleChargeTotal() {
            return engineerDismantleChargeTotal;
        }

        private double getEngineerOtherChargeTotal() {
            return engineerOtherChargeTotal;
        }

        private double getEngineerInsuranceChargeTotal() {
            return engineerInsuranceChargeTotal;
        }

        private double getEngineerTimelinessChargeTotal() {
            return engineerTimelinessChargeTotal;
        }

        private double getEngineerCustomerTimelinessChargeTotal() {
            return engineerCustomerTimelinessChargeTotal;
        }

        private double getEngineerUrgentChargeTotal() {
            return engineerUrgentChargeTotal;
        }

        private double getEngineerTaxFeeTotal() {
            return engineerTaxFeeTotal;
        }

        private double getEngineerInfoFeeTotal() {
            return engineerInfoFeeTotal;
        }

        private double getEngineerDepositTotal() {
            return engineerDepositTotal;
        }

        private double getEngineerPraiseFeeTotal() {
            return engineerPraiseFeeTotal;
        }

        private double getCustomerServiceChargeTotal() {
            return customerServiceChargeTotal;
        }

        private double getCustomerPraiseFeeTotal() {
            return customerPraiseFeeTotal;
        }
        private double getTotalInCharge() {
            return totalInCharge;
        }

        private double getTotalOutCharge() {
            return totalOutCharge;
        }

        private double getCustomerMaterialChargeTotal() { return customerMaterialChargeTotal; }

        private double getCustomerTravelChargeTotal() { return customerTravelChargeTotal; }

        private double getCustomerDismantleChargeTotal() { return customerDismantleChargeTotal; }

        private double getCustomerOtherChargeTotal() { return customerOtherChargeTotal; }

        private double getCustomerTimelinessChargeTotal() { return customerTimelinessChargeTotal; }

        private double getCustomerUrgentChargeTotal() { return customerUrgentChargeTotal; }

        private double getCustomerChargeSumTotal() { return customerChargeSumTotal; }

        private int getRowNumber() {
            return rowNumber;
        }

        private writeCompletionDataExport invoke() {
            for (RPTCompletedOrderDetailsEntity orderMaster : list) {
                rowNumber++;
                int rowSpan = orderMaster.getMaxRow() - 1;
                List<RPTOrderItem> itemList = orderMaster.getItems();
                List<RPTOrderDetail> detailList = orderMaster.getDetails();

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getName()));
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShop().getLabel()));


                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));

                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                RPTOrderDetail detail = null;
                if (detailList != null && detailList.size() > 0) {
                    detail = detailList.get(0);
                }
                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getServicePointNo()));
                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));
                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getContactInfo1()));
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" :
                                detail.getServicePoint().getPaymentType() == null ? "" : detail.getServicePoint().getPaymentType().getLabel()));

                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));

                int qty = (detail == null ? 0 : detail.getQty());
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                totalActualCount = totalActualCount + qty;

                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getRemarks()));

                double engineerServiceCharge = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerServiceCharge);
                engineerServiceChargeTotal = engineerServiceChargeTotal + engineerServiceCharge;

                double engineerMaterialCharge = (detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge());
                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerMaterialCharge);
                engineerMaterialChargeTotal = engineerMaterialChargeTotal + engineerMaterialCharge;

                double engineerTravelCharge = (detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge());
                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTravelCharge);
                engineerTravelChargeTotal = engineerTravelChargeTotal + engineerTravelCharge;

                double engineerDismantleCharge = (detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge());
                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDismantleCharge);
                engineerDismantleChargeTotal = engineerDismantleChargeTotal + engineerDismantleCharge;

                double engineerOtherCharge = (detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge());
                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerOtherCharge);
                engineerOtherChargeTotal = engineerOtherChargeTotal + engineerOtherCharge;

                double engineerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerTimelinessCharge() == null ? 0.0d : detail.getEngineerTimelinessCharge());
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTimelinessCharge);
                engineerTimelinessChargeTotal = engineerTimelinessChargeTotal + engineerTimelinessCharge;

                double engineerCustomerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge() == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge());
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerCustomerTimelinessCharge);
                engineerCustomerTimelinessChargeTotal = engineerCustomerTimelinessChargeTotal + engineerCustomerTimelinessCharge;

                double engineerUrgentCharge = (detail == null ? 0.0d : detail.getEngineerUrgentCharge() == null ? 0.0d : detail.getEngineerUrgentCharge());
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerUrgentCharge);
                engineerUrgentChargeTotal = engineerUrgentChargeTotal + engineerUrgentCharge;

                double engineerChargeSum = engineerServiceCharge + engineerMaterialCharge + engineerTravelCharge + engineerDismantleCharge + engineerOtherCharge + engineerTimelinessCharge+ engineerCustomerTimelinessCharge + engineerUrgentCharge;
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerChargeSum);
                engineerChargeTotal = engineerChargeTotal + engineerChargeSum;

                double engineerInsuranceCharge = (detail == null ? 0.0d : detail.getEngineerInsuranceCharge() == null ? 0.0d : detail.getEngineerInsuranceCharge());
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInsuranceCharge);
                engineerInsuranceChargeTotal = engineerInsuranceChargeTotal + engineerInsuranceCharge;

                double engineerTaxFee = (detail == null ? 0.0d : detail.getEngineerTaxFee() == null ? 0.0d : detail.getEngineerTaxFee());
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTaxFee);
                engineerTaxFeeTotal = engineerTaxFeeTotal + engineerTaxFee;

                double engineerInfoFee = (detail == null ? 0.0d : detail.getEngineerInfoFee() == null ? 0.0d : detail.getEngineerInfoFee());
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInfoFee);
                engineerInfoFeeTotal = engineerInfoFeeTotal + engineerInfoFee;

                double engineerDeposit = (detail == null ? 0.0d : detail.getEngineerDeposit() == null ? 0.0d : detail.getEngineerDeposit());
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);
                engineerDepositTotal = engineerDepositTotal + engineerDeposit;

                double engineerPraiseFee = (detail == null ? 0.0d : detail.getEngineerPraiseFee() == null ? 0.0d : detail.getEngineerPraiseFee());
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFee);
                engineerPraiseFeeTotal = engineerPraiseFeeTotal + engineerPraiseFee;

                double engineerTotalCharge = engineerChargeSum +  engineerInsuranceCharge + engineerTaxFee + engineerInfoFee + engineerDeposit + engineerPraiseFee;
                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTotalCharge);

                totalOutCharge = totalOutCharge + engineerTotalCharge;

                double customerServiceCharge = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerServiceCharge);
                customerServiceChargeTotal = customerServiceChargeTotal+ customerServiceCharge;

                double customerMaterialCharge = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerMaterialCharge);
                customerMaterialChargeTotal = customerMaterialChargeTotal+ customerMaterialCharge;

                double customerTravelCharge = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTravelCharge);
                customerTravelChargeTotal = customerTravelChargeTotal+ customerTravelCharge;

                double customerDismantleCharge = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerDismantleCharge);
                customerDismantleChargeTotal = customerDismantleChargeTotal+ customerDismantleCharge;

                double customerOtherCharge = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerOtherCharge);
                customerOtherChargeTotal = customerOtherChargeTotal+ customerOtherCharge;

                ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTimelinessCharge());
                customerTimelinessChargeTotal = customerTimelinessChargeTotal + orderMaster.getCustomerTimelinessCharge();

                ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerUrgentCharge());
                customerUrgentChargeTotal = customerUrgentChargeTotal + orderMaster.getCustomerUrgentCharge();

                double customerChargeSum = customerServiceCharge + customerMaterialCharge + customerTravelCharge + customerDismantleCharge + customerOtherCharge +orderMaster.getCustomerTimelinessCharge() + orderMaster.getCustomerUrgentCharge() ;
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerChargeSum);
                customerChargeSumTotal = customerChargeSumTotal + customerChargeSum;

                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerPraiseFee());
                customerPraiseFeeTotal = customerPraiseFeeTotal + orderMaster.getCustomerPraiseFee();

                ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTotalCharge());
                totalInCharge = (totalInCharge + orderMaster.getCustomerTotalCharge());

                ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());

                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);

                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getName()));
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShop().getLabel()));
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));

                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));

                        if (detailList != null && index < detailList.size()) {
                            detail = detailList.get(index);

                            ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getServicePointNo()));
                            ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));
                            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getContactInfo1()));
                            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getPaymentType() == null ? "" : detail.getServicePoint().getPaymentType().getLabel()));
                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getServiceTimes());
                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail.getProduct() == null ? "" : detail.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getProductSpec());
                            ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getBrand());

                            int orderDetailQty = (detail == null ? 0 : detail.getQty());
                            ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderDetailQty);
                            totalActualCount = totalActualCount + orderDetailQty;

                            ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getRemarks());

                            engineerServiceCharge = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
                            ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerServiceCharge);
                            engineerServiceChargeTotal = engineerServiceChargeTotal + engineerServiceCharge;

                            engineerMaterialCharge = detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge();
                            ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerMaterialCharge);
                            engineerMaterialChargeTotal = engineerMaterialChargeTotal + engineerMaterialCharge;

                            engineerTravelCharge = detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge();
                            ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTravelCharge);
                            engineerTravelChargeTotal = engineerTravelChargeTotal + engineerTravelCharge;

                            engineerDismantleCharge = detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge();
                            ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDismantleCharge);
                            engineerDismantleChargeTotal = engineerDismantleChargeTotal + engineerDismantleCharge;

                            engineerOtherCharge = detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge();
                            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerOtherCharge);
                            engineerOtherChargeTotal = engineerOtherChargeTotal + engineerOtherCharge;

                            engineerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerTimelinessCharge() == null ? 0.0d : detail.getEngineerTimelinessCharge());
                            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTimelinessCharge);
                            engineerTimelinessChargeTotal = engineerTimelinessChargeTotal + engineerTimelinessCharge;

                            engineerCustomerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge() == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge());
                            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerCustomerTimelinessCharge);
                            engineerCustomerTimelinessChargeTotal = engineerCustomerTimelinessChargeTotal + engineerCustomerTimelinessCharge;

                            engineerUrgentCharge = (detail == null ? 0.0d : detail.getEngineerUrgentCharge() == null ? 0.0d : detail.getEngineerUrgentCharge());
                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerUrgentCharge);
                            engineerUrgentChargeTotal = engineerUrgentChargeTotal + engineerUrgentCharge;

                            engineerChargeSum = engineerServiceCharge + engineerMaterialCharge + engineerTravelCharge + engineerDismantleCharge + engineerOtherCharge + engineerTimelinessCharge +engineerCustomerTimelinessCharge+engineerUrgentCharge;
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerChargeSum);
                            engineerChargeTotal = engineerChargeTotal + engineerChargeSum;

                            engineerInsuranceCharge = (detail == null ? 0.0d : detail.getEngineerInsuranceCharge() == null ? 0.0d : detail.getEngineerInsuranceCharge());
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInsuranceCharge);
                            engineerInsuranceChargeTotal = engineerInsuranceChargeTotal + engineerInsuranceCharge;

                            engineerTaxFee = (detail == null ? 0.0d : detail.getEngineerTaxFee() == null ? 0.0d : detail.getEngineerTaxFee());
                            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTaxFee);
                            engineerTaxFeeTotal = engineerTaxFeeTotal + engineerTaxFee;

                            engineerInfoFee = (detail == null ? 0.0d : detail.getEngineerInfoFee() == null ? 0.0d : detail.getEngineerInfoFee());
                            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInfoFee);
                            engineerInfoFeeTotal = engineerInfoFeeTotal + engineerInfoFee;

                            engineerDeposit = (detail == null ? 0.0d : detail.getEngineerDeposit() == null ? 0.0d : detail.getEngineerDeposit());
                            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);
                            engineerDepositTotal = engineerDepositTotal + engineerDeposit;

                            engineerPraiseFee = (detail == null ? 0.0d : detail.getEngineerPraiseFee() == null ? 0.0d : detail.getEngineerPraiseFee());
                            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFee);
                            engineerPraiseFeeTotal = engineerPraiseFeeTotal + engineerPraiseFee;

                            engineerTotalCharge = engineerChargeSum +  engineerInsuranceCharge + engineerTaxFee + engineerInfoFee + engineerDeposit + engineerPraiseFee;
                            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTotalCharge);

                            totalOutCharge = totalOutCharge + engineerTotalCharge;

                            customerServiceCharge = detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge();
                            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerServiceCharge);
                            customerServiceChargeTotal = customerServiceChargeTotal+ customerServiceCharge;

                            customerMaterialCharge = detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge();
                            ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerMaterialCharge);
                            customerMaterialChargeTotal = customerMaterialChargeTotal+ customerMaterialCharge;

                            customerTravelCharge = detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge();
                            ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTravelCharge);
                            customerTravelChargeTotal = customerTravelChargeTotal+ customerTravelCharge;

                            customerDismantleCharge = detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge();
                            ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerDismantleCharge);
                            customerDismantleChargeTotal = customerDismantleChargeTotal+ customerDismantleCharge;

                            customerOtherCharge = detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge();
                            ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerOtherCharge);
                            customerOtherChargeTotal = customerOtherChargeTotal+ customerOtherCharge;

                            ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);

                            ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);

                            customerChargeSum = customerServiceCharge + customerMaterialCharge + customerTravelCharge + customerDismantleCharge + customerOtherCharge;
                            customerChargeSumTotal = customerChargeSumTotal + customerChargeSum;
                            ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerChargeSum);
                            ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);
                            ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);

                        } else {
                            for (int columnIndex = 6; columnIndex <= 8; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }

                            for (int columnIndex = 10; columnIndex <= 41; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }
                        ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());
                    }
                }
            }
            return this;
        }
    }
}
