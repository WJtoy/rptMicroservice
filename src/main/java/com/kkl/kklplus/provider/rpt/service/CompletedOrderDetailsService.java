package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderDetailsEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.common.WarrantyStatusEnum;
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
import java.util.*;
import java.util.stream.Collectors;


@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CompletedOrderDetailsService extends RptBaseService {

    @Resource
    private CompletedOrderDetailsRptMapper completedOrderDetailsRptMapper;

    //获取报表显示数据
    public Page<RPTCompletedOrderDetailsEntity> getCompletedOrderDetailsRptList(RPTCompletedOrderDetailsSearch search) {
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
                item.setCreateDate(new Date(item.getCreateDt()));
                item.setCloseDate(new Date(item.getCloseDt()));
                item.setChargeDate(new Date(item.getChargeDt()));

                if(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()) != null){
                    item.setWarrantyName(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()).name);
                }

                if (item.getPlanDt() != 0) {
                    item.setPlanDate(new Date(item.getPlanDt()));
                }
                if (item.getAppointmentDt() != 0) {
                    item.setAppointmentDate(new Date(item.getAppointmentDt()));
                }
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
                item.setCreateDate(new Date(item.getCreateDt()));
                if(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()) != null){
                    item.setWarrantyName(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()).name);
                }

                if (item.getCustomerApproveDt() != 0) {
                    item.setCustomerApproveDate(new Date(item.getCustomerApproveDt()));
                }
                if (item.getPlanDt() != 0) {
                    item.setPlanDate(new Date(item.getPlanDt()));
                }

                if (item.getAppointmentDt() != 0) {
                    item.setAppointmentDate(new Date(item.getAppointmentDt()));
                }
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
                item.setCreateDate(new Date(item.getCreateDt()));
                item.setCloseDate(new Date(item.getCloseDt()));
                item.setChargeDate(new Date(item.getChargeDt()));

                if(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()) != null){
                    item.setWarrantyName(WarrantyStatusEnum.valueOf(item.getWarrantyStatus()).name);
                }

                if (item.getPlanDt() != 0) {
                    item.setPlanDate(new Date(item.getPlanDt()));
                }
                if (item.getCustomer().getContractDt() != 0) {
                    item.getCustomer().setContractDate(new Date(item.getCustomer().getContractDt()));
                }
                if (item.getAppointmentDt() != 0) {
                    item.setAppointmentDate(new Date(item.getAppointmentDt()));
                }
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

    public SXSSFWorkbook completedOrderDetailsExport(String searchConditionJson, String reportTitle) {
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 73));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 5));

            ExportExcel.createCell(headerFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "跟进业务员");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 6, 6));

            ExportExcel.createCell(headerFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约业务员");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 7, 7));

            ExportExcel.createCell(headerFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 8, 20));

            ExportExcel.createCell(headerFirstRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 21, 21));

            ExportExcel.createCell(headerFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 22, 22));

            ExportExcel.createCell(headerFirstRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 23, 23));

            ExportExcel.createCell(headerFirstRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "安维人员信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 24, 30));

            ExportExcel.createCell(headerFirstRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "预约上门时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 31, 31));

            ExportExcel.createCell(headerFirstRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "跟综进度");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 32, 32));

            ExportExcel.createCell(headerFirstRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 33, 33));

            ExportExcel.createCell(headerFirstRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际服务项目");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 34, 40));

            ExportExcel.createCell(headerFirstRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付安维费用");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 41, 55));

            ExportExcel.createCell(headerFirstRow, 56, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收客户货款");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 56, 65));

            ExportExcel.createCell(headerFirstRow, 66, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 66, 66));

            ExportExcel.createCell(headerFirstRow, 67, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 67, 67));

            ExportExcel.createCell(headerFirstRow, 68, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工结果");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 68, 68));

            ExportExcel.createCell(headerFirstRow, 69, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "审单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 69, 69));

            ExportExcel.createCell(headerFirstRow, 70, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "退补描述");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 70, 70));

            ExportExcel.createCell(headerFirstRow, 71, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结账日期");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 71, 71));

            ExportExcel.createCell(headerFirstRow, 72, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款日期");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 72, 72));

            ExportExcel.createCell(headerFirstRow, 73, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款描述");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 73, 73));

            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保类型");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完工条码");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "姓名");
            ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "电话");
            ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "支行");
            ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "账户名");
            ExportExcel.createCell(headerSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "账号");
            ExportExcel.createCell(headerSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headerSecondRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
            ExportExcel.createCell(headerSecondRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");
            ExportExcel.createCell(headerSecondRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
            ExportExcel.createCell(headerSecondRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
            ExportExcel.createCell(headerSecondRow, 43, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
            ExportExcel.createCell(headerSecondRow, 44, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
            ExportExcel.createCell(headerSecondRow, 45, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
            ExportExcel.createCell(headerSecondRow, 46, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
            ExportExcel.createCell(headerSecondRow, 47, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "互助基金");
            ExportExcel.createCell(headerSecondRow, 48, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效奖励");
            ExportExcel.createCell(headerSecondRow, 49, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
            ExportExcel.createCell(headerSecondRow, 50, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headerSecondRow, 51, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");
            ExportExcel.createCell(headerSecondRow, 52, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台费");
            ExportExcel.createCell(headerSecondRow, 53, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保金额");
            ExportExcel.createCell(headerSecondRow, 54, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headerSecondRow, 55, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付合计");
            ExportExcel.createCell(headerSecondRow, 56, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
            ExportExcel.createCell(headerSecondRow, 57, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
            ExportExcel.createCell(headerSecondRow, 58, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费");
            ExportExcel.createCell(headerSecondRow, 59, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
            ExportExcel.createCell(headerSecondRow, 60, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");
            ExportExcel.createCell(headerSecondRow, 61, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "汇总");
            ExportExcel.createCell(headerSecondRow, 62, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
            ExportExcel.createCell(headerSecondRow, 63, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headerSecondRow, 64, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headerSecondRow, 65, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收合计");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================

            int totalCount = 0;
            double totalBlocked = 0d;
            int totalActualCount = 0;
            double totalInCharge = 0d;
            double totalOutCharge = 0d;
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
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalBlocked, totalActualCount, totalInCharge, totalOutCharge, rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount =  dataExport.getTotalCount();
                    totalBlocked = dataExport.getTotalBlocked();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge =  dataExport.getTotalInCharge();
                    totalOutCharge =  dataExport.getTotalOutCharge();
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
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalBlocked, totalActualCount, totalInCharge, totalOutCharge, rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount =  dataExport.getTotalCount();
                    totalBlocked =  dataExport.getTotalBlocked();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge =  dataExport.getTotalInCharge();
                    totalOutCharge =  dataExport.getTotalOutCharge();
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
                    writeCompletionDataExport dataExport = new writeCompletionDataExport(xSheet, xStyle, rowIndex, totalCount, totalBlocked, totalActualCount, totalInCharge, totalOutCharge, rowNumber, list).invoke();
                    rowIndex = dataExport.getRowIndex();
                    totalCount = dataExport.getTotalCount();
                    totalBlocked = dataExport.getTotalBlocked();
                    totalActualCount =  dataExport.getTotalActualCount();
                    totalInCharge = dataExport.getTotalInCharge();
                    totalOutCharge = dataExport.getTotalOutCharge();
                    rowNumber = dataExport.getRowNumber();
                    listSize = list.size();
                }
                serviceWrite = serviceWrite + listSize;
                beginLimit++;
            } while (listSize == endLimit);

                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 6));

                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单数量");
                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, finishQty);
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户退补单量");
                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerWrite);
                ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点退补单量");
                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, serviceWrite);
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 16, 38));
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalActualCount);

                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 40, 53));

                ExportExcel.createCell(dataRow, 54, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应付合计");
                ExportExcel.createCell(dataRow, 55, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOutCharge);

                ExportExcel.createCell(dataRow, 56, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 56, 63));

                ExportExcel.createCell(dataRow, 64, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "应收合计");
                ExportExcel.createCell(dataRow, 65, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalInCharge);

                ExportExcel.createCell(dataRow, 66, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 66, 73));


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
        private double totalBlocked;
        private int totalActualCount;
        private double totalInCharge;
        private double totalOutCharge;
        private int rowNumber;
        private Page<RPTCompletedOrderDetailsEntity> list;

        private writeCompletionDataExport(Sheet xSheet, Map<String, CellStyle> xStyle, int rowIndex, int totalCount, double totalBlocked, int totalActualCount, double totalInCharge, double totalOutCharge, int rowNumber, Page<RPTCompletedOrderDetailsEntity> list) {
            this.xSheet = xSheet;
            this.xStyle = xStyle;
            this.rowIndex = rowIndex;
            this.totalCount = totalCount;
            this.totalBlocked = totalBlocked;
            this.totalActualCount = totalActualCount;
            this.totalInCharge = totalInCharge;
            this.totalOutCharge = totalOutCharge;
            this.rowNumber = rowNumber;
            this.list = list;
        }

        private int getRowIndex() {
            return rowIndex;
        }

        private int getTotalCount() {
            return totalCount;
        }

        private double getTotalBlocked() {
            return totalBlocked;
        }

        private int getTotalActualCount() {
            return totalActualCount;
        }

        private double getTotalInCharge() {
            return totalInCharge;
        }

        private double getTotalOutCharge() {
            return totalOutCharge;
        }

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
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getCode()));
                ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getName()));
                ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShop().getLabel()));
                ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getCustomer().getContractDate(), "yyyy-MM-dd")));
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (orderMaster.getPaymentType() == null ? "" : orderMaster.getPaymentType().getLabel()));
                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getSales().getName()));
                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getSales().getName()));
                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getWarrantyName()));


                if (itemList != null && itemList.size() > 0) {
                    RPTOrderItem item = itemList.get(0);
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.joinStringList(item.getUnitBarCodes()));
                    totalCount = totalCount + (item.getQty() == null ? 0 : item.getQty());
                } else {
                    for (int columnIndex = 10; columnIndex <= 15; columnIndex++) {
                        ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    }
                }

                double expectCharge = (orderMaster.getExpectCharge());
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                totalBlocked = totalBlocked + expectCharge;
                ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss")));
                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (orderMaster.getKeFu() == null ? "" : orderMaster.getKeFu().getName()));
                ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (DateUtils.formatDate(orderMaster.getPlanDate(), "yyyy-MM-dd HH:mm:ss")));

                RPTOrderDetail detail = null;
                if (detailList != null && detailList.size() > 0) {
                    detail = detailList.get(0);
                }
                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getServicePointNo()));
                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getContactInfo1()));
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" :
                                detail.getServicePoint().getBank() == null ? "" : detail.getServicePoint().getBank().getLabel()));
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getBankOwner()));
                ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getBankNo()));
                ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" :
                                detail.getServicePoint().getPaymentType() == null ? "" : detail.getServicePoint().getPaymentType().getLabel()));

                ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        DateUtils.formatDate(orderMaster.getAppointmentDate(), "yyyy-MM-dd HH:mm:ss"));
                ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTrackingComment());

                ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss")));

                ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? 0 : detail.getServiceTimes()));
                ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getProductSpec()));
                ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getBrand()));

                int qty = (detail == null ? 0 : detail.getQty());
                ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                totalActualCount = totalActualCount + qty;

                ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getRemarks()));

                double engineerServiceCharge = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
                ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerServiceCharge);

                double engineerMaterialCharge = (detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge());
                ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerMaterialCharge);

                double engineerTravelCharge = (detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge());
                ExportExcel.createCell(dataRow, 43, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTravelCharge);

                double engineerDismantleCharge = (detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge());
                ExportExcel.createCell(dataRow, 44, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDismantleCharge);

                double engineerOtherCharge = (detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge());
                ExportExcel.createCell(dataRow, 45, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerOtherCharge);

                double engineerChargeSum = engineerServiceCharge + engineerMaterialCharge + engineerTravelCharge + engineerDismantleCharge + engineerOtherCharge;
                ExportExcel.createCell(dataRow, 46, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerChargeSum);

                double engineerInsuranceCharge = (detail == null ? 0.0d : detail.getEngineerInsuranceCharge() == null ? 0.0d : detail.getEngineerInsuranceCharge());
                ExportExcel.createCell(dataRow, 47, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInsuranceCharge);

                double engineerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerTimelinessCharge() == null ? 0.0d : detail.getEngineerTimelinessCharge());
                ExportExcel.createCell(dataRow, 48, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTimelinessCharge);

                double engineerCustomerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge() == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge());
                ExportExcel.createCell(dataRow, 49, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerCustomerTimelinessCharge);

                double engineerUrgentCharge = (detail == null ? 0.0d : detail.getEngineerUrgentCharge() == null ? 0.0d : detail.getEngineerUrgentCharge());
                ExportExcel.createCell(dataRow, 50, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerUrgentCharge);

                double engineerTaxFee = (detail == null ? 0.0d : detail.getEngineerTaxFee() == null ? 0.0d : detail.getEngineerTaxFee());
                ExportExcel.createCell(dataRow, 51, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTaxFee);

                double engineerInfoFee = (detail == null ? 0.0d : detail.getEngineerInfoFee() == null ? 0.0d : detail.getEngineerInfoFee());
                ExportExcel.createCell(dataRow, 52, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInfoFee);

                double engineerDeposit = (detail == null ? 0.0d : detail.getEngineerDeposit() == null ? 0.0d : detail.getEngineerDeposit());
                ExportExcel.createCell(dataRow, 53, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);

                double engineerPraiseFee = (detail == null ? 0.0d : detail.getEngineerPraiseFee() == null ? 0.0d : detail.getEngineerPraiseFee());
                ExportExcel.createCell(dataRow, 54, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFee);
                double engineerTotalCharge = engineerChargeSum +  engineerInsuranceCharge + engineerTimelinessCharge + engineerCustomerTimelinessCharge + engineerUrgentCharge + engineerTaxFee + engineerInfoFee + engineerDeposit + engineerPraiseFee;
                ExportExcel.createCell(dataRow, 55, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTotalCharge);

                totalOutCharge = totalOutCharge + engineerTotalCharge;

                double customerServiceCharge = (detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge());
                ExportExcel.createCell(dataRow, 56, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerServiceCharge);

                double customerMaterialCharge = (detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge());
                ExportExcel.createCell(dataRow, 57, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerMaterialCharge);

                double customerTravelCharge = (detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge());
                ExportExcel.createCell(dataRow, 58, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTravelCharge);

                double customerDismantleCharge = (detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge());
                ExportExcel.createCell(dataRow, 59, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerDismantleCharge);

                double customerOtherCharge = (detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge());
                ExportExcel.createCell(dataRow, 60, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerOtherCharge);

                double customerChargeSum = customerServiceCharge + customerMaterialCharge + customerTravelCharge + customerDismantleCharge + customerOtherCharge;
                ExportExcel.createCell(dataRow, 61, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerChargeSum);

                ExportExcel.createCell(dataRow, 62, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTimelinessCharge());
                ExportExcel.createCell(dataRow, 63, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerUrgentCharge());
                ExportExcel.createCell(dataRow, 64, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerPraiseFee());
                ExportExcel.createCell(dataRow, 65, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTotalCharge());
                totalInCharge = (totalInCharge + orderMaster.getCustomerTotalCharge());


                ExportExcel.createCell(dataRow, 66, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTimeliness());
                ExportExcel.createCell(dataRow, 67, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel());
                ExportExcel.createCell(dataRow, 68, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getAppCompleteType() == null ? "" : orderMaster.getAppCompleteType().getLabel());
                ExportExcel.createCell(dataRow, 69, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getChargeDate(), "yyyy-MM-dd HH:mm:ss"));
                ExportExcel.createCell(dataRow, 70, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());
                ExportExcel.createCell(dataRow, 71, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCustomerInvoiceDate(), "yyyy-MM-dd HH:mm:ss"));

                ExportExcel.createCell(dataRow, 72, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : DateUtils.formatDate(orderMaster.getEngineerInvoiceDate(), "yyyy-MM-dd HH:mm:ss"));
                ExportExcel.createCell(dataRow, 73, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail == null ? "" : orderMaster.getEngineerInvoiceRemarks());

                if (rowSpan > 0) {
                    for (int index = 1; index <= rowSpan; index++) {
                        dataRow = xSheet.createRow(rowIndex++);

                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getCode()));
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getName()));
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getShop().getLabel()));
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getCustomer().getContractDate(), "yyyy-MM-dd")));
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (orderMaster.getPaymentType() == null ? "" : orderMaster.getPaymentType().getLabel()));
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getSales().getName()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomer().getSales().getName()));
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getWarrantyName()));

                        if (itemList != null && index < itemList.size()) {
                            RPTOrderItem item = itemList.get(index);

                            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                            ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                            ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUnitBarcode());
                            totalCount = totalCount + StringUtils.toInteger(item.getQty());
                        } else {
                            for (int columnIndex = 10; columnIndex <= 15; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                        expectCharge = (orderMaster.getExpectCharge());
                        ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, expectCharge);
                        totalBlocked = totalBlocked + expectCharge;
                        ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getDescription()));
                        ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserName()));
                        ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserPhone()));
                        ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                        ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (DateUtils.formatDate(orderMaster.getCustomerApproveDate(), "yyyy-MM-dd HH:mm:ss")));
                        ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (orderMaster.getKeFu() == null ? "" : orderMaster.getKeFu().getName()));
                        ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (DateUtils.formatDate(orderMaster.getPlanDate(), "yyyy-MM-dd HH:mm:ss")));

                        if (detailList != null && index < detailList.size()) {
                            detail = detailList.get(index);

                            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getServicePointNo()));
                            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));
                            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getContactInfo1()));
                            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getBank() == null ? "" : detail.getServicePoint().getBank().getLabel()));
                            ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getBankOwner()));
                            ExportExcel.createCell(dataRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getBankNo()));
                            ExportExcel.createCell(dataRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (detail == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint() == null ? "" : detail.getServicePoint().getPaymentType() == null ? "" : detail.getServicePoint().getPaymentType().getLabel()));
                            ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (DateUtils.formatDate(orderMaster.getAppointmentDate(), "yyyy-MM-dd HH:mm:ss")));
                            ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");//cell.setCellValue(orderMaster.get("trackingComment") == null ? "" : orderMaster.get("trackingComment").toString());
                            ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                    (DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss")));
                            ExportExcel.createCell(dataRow, 34, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getServiceTimes());
                            ExportExcel.createCell(dataRow, 35, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                            ExportExcel.createCell(dataRow, 36, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail.getProduct() == null ? "" : detail.getProduct().getName()));
                            ExportExcel.createCell(dataRow, 37, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getProductSpec());
                            ExportExcel.createCell(dataRow, 38, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getBrand());

                            int orderDetailQty = (detail == null ? 0 : detail.getQty());
                            ExportExcel.createCell(dataRow, 39, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderDetailQty);
                            totalActualCount = totalActualCount + orderDetailQty;

                            ExportExcel.createCell(dataRow, 40, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, detail.getRemarks());

                            engineerServiceCharge = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
                            ExportExcel.createCell(dataRow, 41, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerServiceCharge);

                            engineerMaterialCharge = detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge();
                            ExportExcel.createCell(dataRow, 42, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerMaterialCharge);

                            engineerTravelCharge = detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge();
                            ExportExcel.createCell(dataRow, 43, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTravelCharge);

                            engineerDismantleCharge = detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge();
                            ExportExcel.createCell(dataRow, 44, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDismantleCharge);

                            engineerOtherCharge = detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge();
                            ExportExcel.createCell(dataRow, 45, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerOtherCharge);
                            engineerChargeSum = engineerServiceCharge + engineerMaterialCharge + engineerTravelCharge + engineerDismantleCharge + engineerOtherCharge;
                            ExportExcel.createCell(dataRow, 46, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerChargeSum);

                            engineerInsuranceCharge = (detail == null ? 0.0d : detail.getEngineerInsuranceCharge() == null ? 0.0d : detail.getEngineerInsuranceCharge());
                            ExportExcel.createCell(dataRow, 47, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInsuranceCharge);

                            engineerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerTimelinessCharge() == null ? 0.0d : detail.getEngineerTimelinessCharge());
                            ExportExcel.createCell(dataRow, 48, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTimelinessCharge);

                            engineerCustomerTimelinessCharge = (detail == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge() == null ? 0.0d : detail.getEngineerCustomerTimelinessCharge());
                            ExportExcel.createCell(dataRow, 49, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerCustomerTimelinessCharge);

                            engineerUrgentCharge = (detail == null ? 0.0d : detail.getEngineerUrgentCharge() == null ? 0.0d : detail.getEngineerUrgentCharge());
                            ExportExcel.createCell(dataRow, 50, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerUrgentCharge);

                            engineerTaxFee = (detail == null ? 0.0d : detail.getEngineerTaxFee() == null ? 0.0d : detail.getEngineerTaxFee());
                            ExportExcel.createCell(dataRow, 51, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTaxFee);

                            engineerInfoFee = (detail == null ? 0.0d : detail.getEngineerInfoFee() == null ? 0.0d : detail.getEngineerInfoFee());
                            ExportExcel.createCell(dataRow, 52, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerInfoFee);

                            engineerDeposit = (detail == null ? 0.0d : detail.getEngineerDeposit() == null ? 0.0d : detail.getEngineerDeposit());
                            ExportExcel.createCell(dataRow, 53, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);

                            engineerPraiseFee = (detail == null ? 0.0d : detail.getEngineerPraiseFee() == null ? 0.0d : detail.getEngineerPraiseFee());
                            ExportExcel.createCell(dataRow, 54, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFee);
                            engineerTotalCharge = engineerChargeSum +  engineerInsuranceCharge + engineerTimelinessCharge + engineerCustomerTimelinessCharge + engineerUrgentCharge + engineerTaxFee + engineerInfoFee + engineerDeposit + engineerPraiseFee;
                            ExportExcel.createCell(dataRow, 55, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerTotalCharge);

                            totalOutCharge = totalOutCharge + engineerTotalCharge;

                            customerServiceCharge = detail == null ? 0.0d : detail.getServiceCharge() == null ? 0.0d : detail.getServiceCharge();
                            ExportExcel.createCell(dataRow, 56, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerServiceCharge);

                            customerMaterialCharge = detail == null ? 0.0d : detail.getMaterialCharge() == null ? 0.0d : detail.getMaterialCharge();
                            ExportExcel.createCell(dataRow, 57, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerMaterialCharge);

                            customerTravelCharge = detail == null ? 0.0d : detail.getTravelCharge() == null ? 0.0d : detail.getTravelCharge();
                            ExportExcel.createCell(dataRow, 58, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerTravelCharge);

                            customerDismantleCharge = detail == null ? 0.0d : detail.getExpressCharge() == null ? 0.0d : detail.getExpressCharge();
                            ExportExcel.createCell(dataRow, 59, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerDismantleCharge);

                            customerOtherCharge = detail == null ? 0.0d : detail.getOtherCharge() == null ? 0.0d : detail.getOtherCharge();
                            ExportExcel.createCell(dataRow, 60, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerOtherCharge);

                            customerChargeSum = customerServiceCharge + customerMaterialCharge + customerTravelCharge + customerDismantleCharge + customerOtherCharge;
                            ExportExcel.createCell(dataRow, 61, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, customerChargeSum);
                            ExportExcel.createCell(dataRow, 62, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);
                            ExportExcel.createCell(dataRow, 63, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);
                            ExportExcel.createCell(dataRow, 64, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);
                            ExportExcel.createCell(dataRow, 65, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0);

                            ExportExcel.createCell(dataRow, 69, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getEngineerInvoiceDate(), "yyyy-MM-dd HH:mm:ss"));
                            ExportExcel.createCell(dataRow, 70, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getEngineerInvoiceRemarks());
                        } else {
                            for (int columnIndex = 23; columnIndex <= 28; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }

                            for (int columnIndex = 32; columnIndex <= 64; columnIndex++) {
                                ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            }
                        }

                        ExportExcel.createCell(dataRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                DateUtils.formatDate(orderMaster.getAppointmentDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getTrackingComment());

                        ExportExcel.createCell(dataRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (DateUtils.formatDate(orderMaster.getCloseDate(), "yyyy-MM-dd HH:mm:ss")));

                        ExportExcel.createCell(dataRow, 66, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomerTimeliness());
                        ExportExcel.createCell(dataRow, 67, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus() == null ? "" : orderMaster.getStatus().getLabel());
                        ExportExcel.createCell(dataRow, 68, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getAppCompleteType() == null ? "" : orderMaster.getAppCompleteType().getLabel());
                        ExportExcel.createCell(dataRow, 69, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getChargeDate(), "yyyy-MM-dd HH:mm:ss"));
                        ExportExcel.createCell(dataRow, 70, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getWriteOffRemarks());
                        ExportExcel.createCell(dataRow, 71, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCustomerInvoiceDate(), "yyyy-MM-dd HH:mm:ss"));
                    }
                }
            }
            return this;
        }
    }
}
