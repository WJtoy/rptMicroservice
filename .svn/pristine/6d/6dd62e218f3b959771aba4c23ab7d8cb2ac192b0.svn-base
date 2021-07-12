package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTServicePointBalanceEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;

import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointBalanceRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointBalanceRptService extends RptBaseService {

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private ServicePointBalanceRptMapper servicePointBalanceRptMapper;

    @Autowired
    private MSEngineerService msEngineerService;




        public Page<RPTServicePointBalanceEntity> getServicePointBalanceRptData(RPTServicePointWriteOffSearch search){
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginWriteOffCreateDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getBeginWriteOffCreateDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Date preMonthDate = DateUtils.addMonth(queryDate, -1);
        Page<RPTServicePointBalanceEntity> page = new Page<>();
        List<RPTServicePointBalanceEntity> list = Lists.newArrayList();

        if (new Date().getTime() < queryDate.getTime()) {
            return page;
        }
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

            Set<Long> servicePointIdSet = Sets.newHashSet();
            long servicePoindId = 0L;
            if (null != search.getServicePointId()) {
                servicePoindId = search.getServicePointId();
                servicePointIdSet.add(servicePoindId);
            }


            page = servicePointBalanceRptMapper.getServicePointPayableData(servicePoindId, search.getPaymentType(), selectedYear, selectedMonth, search.getProductCategoryIds(),search.getPageNo(),search.getPageSize());

            Map<Long, RPTServicePointBalanceEntity> payablePaidBalanceMap = Maps.newHashMap();
            for (RPTServicePointBalanceEntity item : page) {
                servicePointIdSet.add(item.getServicePointId());
                payablePaidBalanceMap.put(item.getServicePointId(), item);
            }


            List<Long> servicePointIds = Lists.newArrayList(servicePointIdSet);
            List<RPTServicePointBalanceEntity> preBalanceList = servicePointBalanceRptMapper.getServicePointBalanceData((servicePointIds.size() <= 100 ? servicePointIds : null),
                    search.getPaymentType(), DateUtils.getYear(preMonthDate), DateUtils.getMonth(preMonthDate), search.getProductCategoryIds());

            Map<Long, RPTServicePointBalanceEntity> preBalanceListMap = Maps.newHashMap();
            for (RPTServicePointBalanceEntity item : preBalanceList) {
                servicePointIdSet.add(item.getServicePointId());
                preBalanceListMap.put(item.getServicePointId(), item);
            }

            servicePointIds = Lists.newArrayList(servicePointIdSet);//使用最新的网点ID列表
            servicePointIds = servicePointIds.stream().sorted().collect(Collectors.toList());
            RPTServicePointBalanceEntity rptEntity;

            String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType", "primaryId"};
            Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray), null);
            List<Long> engineerIds = servicePointMap.entrySet().stream().map(i -> i.getValue().getPrimaryId()).distinct().collect(Collectors.toList());
            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));
            RPTDict bankDict;
            RPTDict  paymentTypDict;
            RPTDict  paymentTypDic;
            RPTEngineer engineer = new RPTEngineer();
            Map<String, RPTDict> bankMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
            Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);

            for (Long sid : servicePointIds) {
                rptEntity = new RPTServicePointBalanceEntity();
                RPTServicePointBalanceEntity  payableEntity = payablePaidBalanceMap.get(sid);
                MDServicePointViewModel servicePoint = servicePointMap.get(sid);
                RPTServicePointBalanceEntity  balanceEntity = preBalanceListMap.get(sid);

                if (servicePoint != null) {
                    engineer = engineerMap.get(servicePoint.getPrimaryId());
                }

                if (servicePoint != null) {
                    if(balanceEntity !=null){
                        rptEntity.setPreBalance(balanceEntity.getPreBalance());
                    }
                    if(payableEntity !=null){
                        rptEntity.setPaidAmount(payableEntity.getPaidAmount());
                        rptEntity.setTheBalance(payableEntity.getTheBalance());
                        rptEntity.setPayableAmount(payableEntity.getPayableAmount());
                    }
                    rptEntity.setServicePointId(servicePoint.getId());
                    rptEntity.setServicePointNo(servicePoint.getServicePointNo());
                    rptEntity.setServicePointName(servicePoint.getName());
                    paymentTypDict = paymentTypeMap.get(servicePoint.getPaymentType().toString());
                    if (paymentTypDict != null && StringUtils.isNotBlank(paymentTypDict.getLabel())) {
                        rptEntity.setPaymentType(paymentTypDict);
                        if (search.getPaymentType() != null) {
                            paymentTypDic = paymentTypeMap.get(search.getPaymentType().toString());
                            if (!paymentTypDic.getLabel().equals(paymentTypDict.getLabel())) {
                                rptEntity.setPaymentType(null);
                            }
                        }
                    }
                    rptEntity.setEngineerName(engineer == null ? "" : engineer.getName());
                    rptEntity.setContactInfo1(servicePoint.getContactInfo1());
                    rptEntity.setContactInfo2(servicePoint.getContactInfo2());
                    bankDict = bankMap.get(servicePoint.getBank().toString());
                    if (bankDict != null && StringUtils.isNotBlank(bankDict.getLabel())) {
                        rptEntity.setBank(bankDict);
                    }
                    rptEntity.setBankOwner(servicePoint.getBankOwner());
                    rptEntity.setBankNo(servicePoint.getBankNo());

                }
                list.add(rptEntity);
            }

         page.clear();
         page.addAll(list);
         return page;
    }

    public List<RPTServicePointBalanceEntity> getServicePointBalanceReport(RPTServicePointWriteOffSearch search){
        Integer selectedYear = DateUtils.getYear(new Date(search.getBeginWriteOffCreateDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getBeginWriteOffCreateDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        Date preMonthDate = DateUtils.addMonth(queryDate, -1);
        List<RPTServicePointBalanceEntity> list = Lists.newArrayList();

        if (new Date().getTime() < queryDate.getTime()) {
            return list;
        }
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        Set<Long> servicePointIdSet = Sets.newHashSet();
        long servicePoindId = 0L;
        if (null != search.getServicePointId()) {
            servicePoindId = search.getServicePointId();
            servicePointIdSet.add(servicePoindId);
        }


        List<RPTServicePointBalanceEntity> payablePaidBalanceList = servicePointBalanceRptMapper.getServicePointPayableReport(servicePoindId, search.getPaymentType(), selectedYear, selectedMonth, search.getProductCategoryIds());

        Map<Long, RPTServicePointBalanceEntity> payablePaidBalanceMap = Maps.newHashMap();
        for (RPTServicePointBalanceEntity item : payablePaidBalanceList) {
            servicePointIdSet.add(item.getServicePointId());
            payablePaidBalanceMap.put(item.getServicePointId(), item);
        }


        List<Long> servicePointIds = Lists.newArrayList(servicePointIdSet);
        List<RPTServicePointBalanceEntity> preBalanceList = servicePointBalanceRptMapper.getServicePointBalanceData((servicePointIds.size() <= 100 ? servicePointIds : null),
                search.getPaymentType(), DateUtils.getYear(preMonthDate), DateUtils.getMonth(preMonthDate), search.getProductCategoryIds());

        Map<Long, RPTServicePointBalanceEntity> preBalanceListMap = Maps.newHashMap();
        for (RPTServicePointBalanceEntity item : preBalanceList) {
            servicePointIdSet.add(item.getServicePointId());
            preBalanceListMap.put(item.getServicePointId(), item);
        }

        servicePointIds = Lists.newArrayList(servicePointIdSet);//使用最新的网点ID列表
        servicePointIds = servicePointIds.stream().sorted().collect(Collectors.toList());
        RPTServicePointBalanceEntity rptEntity;

        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType", "primaryId"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray), null);
        List<Long> engineerIds = servicePointMap.entrySet().stream().map(i -> i.getValue().getPrimaryId()).distinct().collect(Collectors.toList());
        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));
        RPTDict bankDict;
        RPTDict  paymentTypDict;
        RPTDict  paymentTypDic;
        RPTEngineer engineer = new RPTEngineer();
        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);

        for (Long sid : servicePointIds) {
            rptEntity = new RPTServicePointBalanceEntity();
            MDServicePointViewModel servicePoint = servicePointMap.get(sid);
            RPTServicePointBalanceEntity  balanceEntity = preBalanceListMap.get(sid);
            RPTServicePointBalanceEntity  payableEntity = payablePaidBalanceMap.get(sid);

            if (servicePoint != null) {
                engineer = engineerMap.get(servicePoint.getPrimaryId());
            }

            if (servicePoint != null) {
                if(balanceEntity !=null){
                    rptEntity.setPreBalance(balanceEntity.getPreBalance());
                }
                if(payableEntity !=null){
                    rptEntity.setPaidAmount(payableEntity.getPaidAmount());
                    rptEntity.setTheBalance(payableEntity.getTheBalance());
                    rptEntity.setPayableAmount(payableEntity.getPayableAmount());
                }
                rptEntity.setServicePointId(servicePoint.getId());
                rptEntity.setServicePointNo(servicePoint.getServicePointNo());
                rptEntity.setServicePointName(servicePoint.getName());
                paymentTypDict = paymentTypeMap.get(servicePoint.getPaymentType().toString());
                if (paymentTypDict != null && StringUtils.isNotBlank(paymentTypDict.getLabel())) {
                    rptEntity.setPaymentType(paymentTypDict);
                    if (search.getPaymentType() != null) {
                        paymentTypDic = paymentTypeMap.get(search.getPaymentType().toString());
                        if (!paymentTypDic.getLabel().equals(paymentTypDict.getLabel())) {
                            rptEntity.setPaymentType(null);
                        }
                    }
                }
                rptEntity.setEngineerName(engineer == null ? "" : engineer.getName());
                rptEntity.setContactInfo1(servicePoint.getContactInfo1());
                rptEntity.setContactInfo2(servicePoint.getContactInfo2());
                bankDict = bankMap.get(servicePoint.getBank().toString());
                if (bankDict != null && StringUtils.isNotBlank(bankDict.getLabel())) {
                    rptEntity.setBank(bankDict);
                }
                rptEntity.setBankOwner(servicePoint.getBankOwner());
                rptEntity.setBankNo(servicePoint.getBankNo());

            }
            list.add(rptEntity);
        }
        return list;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
        Set<Long> servicePointIdSet = Sets.newHashSet();
        if (searchCondition.getBeginWriteOffCreateDate() != null && searchCondition.getEndWriteOffCreateDate() != null){
            Integer selectedYear = DateUtils.getYear(new Date(searchCondition.getBeginWriteOffCreateDate()));
            Integer selectedMonth = DateUtils.getMonth(new Date(searchCondition.getBeginWriteOffCreateDate()));
            Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
            Date preMonthDate = DateUtils.addMonth(queryDate, -1);
            if (new Date().getTime() < queryDate.getTime()) {
                result = false;
                return result ;
            }
            long servicePoindId = 0L;
            if(null != searchCondition.getServicePointId()){
                servicePoindId = searchCondition.getServicePointId();
                servicePointIdSet.add(servicePoindId);
            }

            List<Long> servicePointIds = Lists.newArrayList(servicePointIdSet);
            Integer servicePointPayableDataSum = servicePointBalanceRptMapper.getServicePointPayableDataSum(servicePoindId, searchCondition.getPaymentType(), selectedYear, selectedMonth, searchCondition.getProductCategoryIds());
            if(servicePointPayableDataSum != 0 ){
                return  result;
            }
            Integer servicePointBalanceDataSum = servicePointBalanceRptMapper.getServicePointBalanceDataSum(servicePointIds, searchCondition.getPaymentType(), DateUtils.getYear(preMonthDate), DateUtils.getMonth(preMonthDate),searchCondition.getProductCategoryIds());
            result = servicePointBalanceDataSum > 0;
        }
        return result;
    }




    public SXSSFWorkbook servicePointBalanceRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
            List<RPTServicePointBalanceEntity>  list =getServicePointBalanceReport(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "负责人");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手机");
            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话");
            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户银行");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户人");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "银行账号");

            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月余额");
            ExportExcel.createCell(headRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点应付");
            ExportExcel.createCell(headRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点已付");
            ExportExcel.createCell(headRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月余额");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {

                double totalPreBalance = 0.00;
                double totalPayableAmount = 0.00;
                double totalPaidAmount = 0.00;
                double totalTheBalance = 0.00;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTServicePointBalanceEntity rowData = list.get(dataRowIndex);
                    int rowSpan = rowData.getMaxRow() - 1;

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                    }
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                    }
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getContactInfo1());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getContactInfo2());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                    }
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBank() == null ? "" : rowData.getBank().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 6, 6));
                    }
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankOwner());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 7, 7));
                    }
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 8, 8));
                    }

                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,rowData.getPreBalance());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,rowData.getPayableAmount());
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,rowData.getPaidAmount());
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,rowData.getTheBalance());

                    totalPreBalance = totalPreBalance + ( rowData.getPreBalance());
                    totalPayableAmount = totalPayableAmount + (rowData.getPayableAmount());
                    totalPaidAmount = totalPaidAmount + (rowData.getPaidAmount());
                    totalTheBalance = totalTheBalance + (rowData.getTheBalance());
                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 9));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPreBalance);
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPayableAmount);
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPaidAmount);
                ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTheBalance);
            }

        } catch (Exception e) {
            log.error("【ServicePointBalanceRptService.servicePointBalanceRptExport】报表网点余额写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
