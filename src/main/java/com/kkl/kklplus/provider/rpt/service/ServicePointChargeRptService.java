package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTServicePointPaySummaryEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointPaySummarySearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.RPTServicePointChargeEntity;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointChargeRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
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

import javax.annotation.Resource;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointChargeRptService extends RptBaseService {

    @Resource
    private ServicePointChargeRptMapper servicePointChargeRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;

    public Page<RPTServicePointPaySummaryEntity> getServicePointPaySummaryRptData(RPTServicePointPaySummarySearch search) {
        PageHelper.startPage(search.getPageNo(), search.getPageSize());
        Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        Page<RPTServicePointPaySummaryEntity> returnPage;
        search.setSystemId(RptCommonUtils.getSystemId());
        search.setYearMonth(yearMonth);
        search.setQuarter(quarter);

        returnPage = servicePointChargeRptMapper.getServicePointPaySummaryRptData(search);
        Set<Long> servicePointIds = returnPage.stream().map(RPTServicePointPaySummaryEntity::getServicePointId).distinct().collect(Collectors.toSet());

        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType", "primaryId", "address", "remarks", "invoiceFlag", "discountFlag"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        MDServicePointViewModel servicePointVM;

        List<Long> engineerIds = servicePointMap.values().stream().map(MDServicePointViewModel::getPrimaryId).distinct().collect(Collectors.toList());

        Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
        if (engineerIds != null && !engineerIds.isEmpty()) {
            List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Arrays.asList("id", "name", "appFlag"));
            if (engineerList != null && !engineerList.isEmpty()) {
                engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
            }
        }
        RPTEngineer engineer;
        Map<String, RPTDict> bankTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        RPTDict engineerPaymentTypeDict;
        RPTDict bankTypeDict;
        RPTServicePoint servicePoint;
        Double platformFee;
        for (RPTServicePointPaySummaryEntity entity : returnPage) {
            servicePointVM = servicePointMap.get(entity.getServicePointId());
            if (servicePointVM != null) {
                servicePoint = new RPTServicePoint();
                servicePoint.setServicePointNo(servicePointVM.getServicePointNo());
                servicePoint.setName(servicePointVM.getName());
                servicePoint.setContactInfo1(servicePointVM.getContactInfo1());
                servicePoint.setContactInfo2(servicePointVM.getContactInfo2());
                if (servicePointVM.getBank() != null) {
                    bankTypeDict = bankTypeMap.get(servicePointVM.getBank().toString());
                    if (bankTypeDict != null) {
                        servicePoint.setBank(bankTypeDict);
                    }
                }
                servicePoint.setBankOwner(servicePointVM.getBankOwner());
                servicePoint.setBankNo(servicePointVM.getBankNo());
                if (servicePointVM.getPaymentType() != null) {
                    engineerPaymentTypeDict = paymentTypeMap.get(servicePointVM.getPaymentType().toString());
                    if (engineerPaymentTypeDict != null) {
                        servicePoint.setPaymentType(engineerPaymentTypeDict);
                    }
                }
                entity.setServicePoint(servicePoint);
                engineer = engineerMap.get(servicePointVM.getPrimaryId());
                if (engineer != null) {
                    entity.setPrimaryName(engineer.getName());
                }

                entity.setAddress(servicePointVM.getAddress());

                entity.setServicePointRemarks(servicePointVM.getRemarks() == null ? "" : servicePointVM.getRemarks());


                entity.setInvoiceFlag(servicePointVM.getInvoiceFlag() == null ? 0 : servicePointVM.getInvoiceFlag());


                entity.setDiscountFlag(servicePointVM.getDiscountFlag() == null ? 0 : servicePointVM.getDiscountFlag());

            }
            platformFee = entity.getPlatformFee();
            entity.setPaidAmount(entity.getPaidAmount() - platformFee);
        }

        return returnPage;
    }


    public Page<RPTServicePointPaySummaryEntity> getServicePointCostPerRptData(RPTServicePointPaySummarySearch search) {
        PageHelper.startPage(search.getPageNo(), search.getPageSize());
        Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        Page<RPTServicePointPaySummaryEntity> returnPage = new Page<>();
        search.setSystemId(RptCommonUtils.getSystemId());
        search.setYearMonth(yearMonth);
        search.setQuarter(quarter);

        returnPage = servicePointChargeRptMapper.getServicePointCostPerRptData(search);
        Set<Long> servicePointIds = returnPage.stream().map(RPTServicePointPaySummaryEntity::getServicePointId).collect(Collectors.toSet());

        Map<Long,RPTArea> cacheCountyAreaMap =  areaCacheService.getAllCountyMap();
        List<MDServicePointArea> serviceAreaList = msServicePointService.getServicePointAreaByIds(Lists.newArrayList(servicePointIds));
//        List<RPTServicePointPaySummaryEntity> serviceAreaList = servicePointChargeRptMapper.getServicePointServiceAreas(Lists.newArrayList(servicePointIds));
        Map<Long, List<MDServicePointArea>> serviceAreaMap = serviceAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId));

        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "paymentType", "primaryId", "address", "remarks"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        MDServicePointViewModel servicePointVM;

        List<Long> engineerIds = servicePointMap.values().stream().map(MDServicePointViewModel::getPrimaryId).distinct().collect(Collectors.toList());

        Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
        if (engineerIds != null && !engineerIds.isEmpty()) {
            List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Arrays.asList("id", "name", "appFlag"));
            if (engineerList != null && !engineerList.isEmpty()) {
                engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
            }
        }
        RPTEngineer engineer;
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        RPTServicePoint servicePoint;
        RPTDict engineerPaymentTypeDict;
        for (RPTServicePointPaySummaryEntity entity : returnPage) {
            servicePointVM = servicePointMap.get(entity.getServicePointId());
            if (servicePointVM != null) {
                servicePoint = new RPTServicePoint();
                servicePoint.setServicePointNo(servicePointVM.getServicePointNo());
                servicePoint.setName(servicePointVM.getName());
                servicePoint.setContactInfo1(servicePointVM.getContactInfo1());
                servicePoint.setContactInfo2(servicePointVM.getContactInfo2());

                if (servicePointVM.getPaymentType() != null) {
                    engineerPaymentTypeDict = paymentTypeMap.get(servicePointVM.getPaymentType().toString());
                    if (engineerPaymentTypeDict != null) {
                        servicePoint.setPaymentType(engineerPaymentTypeDict);
                    }
                }
                entity.setServicePoint(servicePoint);
                engineer = engineerMap.get(servicePointVM.getPrimaryId());
                if (engineer != null) {
                    entity.setPrimaryName(engineer.getName());
                    entity.setAppFlag(engineer.getAppFlag());
                }
                entity.setAddress(servicePointVM.getAddress());
                StringBuilder sb = new StringBuilder();
                if (serviceAreaMap.get(entity.getServicePointId()) != null) {
                    int index = 0;
                    for(MDServicePointArea item : serviceAreaMap.get(entity.getServicePointId())){
                        RPTArea  rptArea = cacheCountyAreaMap.get(item.getAreaId());
                        if(index != 0){
                            sb.append(",");
                        }
                        if(rptArea != null){
                            index++;
                            sb.append(rptArea.getName());
                        }

                    }
                    entity.setServiceAreas(sb.toString());
                }
                entity.setServicePointRemarks(servicePointVM.getRemarks() == null ? "" : servicePointVM.getRemarks());
            }
        }
        return returnPage;
    }


    public List<RPTServicePointPaySummaryEntity> getServicePointCostPerRptList(RPTServicePointPaySummarySearch search) {
        Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        List<RPTServicePointPaySummaryEntity> returnPage = Lists.newArrayList();
        search.setSystemId(RptCommonUtils.getSystemId());
        search.setYearMonth(yearMonth);
        search.setQuarter(quarter);

        returnPage = servicePointChargeRptMapper.getServicePointCostPerRptList(search);
        Set<Long> servicePointIds = returnPage.stream().map(RPTServicePointPaySummaryEntity::getServicePointId).distinct().collect(Collectors.toSet());

        Map<Long,RPTArea> cacheCountyAreaMap =  areaCacheService.getAllCountyMap();
        List<MDServicePointArea> serviceAreaList;
        if(search.getServicePointId()!=null && search.getServicePointId() != 0){
            serviceAreaList = msServicePointService.getServicePointAreaByIds(Lists.newArrayList(servicePointIds));
        }else {
            serviceAreaList = msServicePointService.getServicePointServiceAreas();
        }
//        List<RPTServicePointPaySummaryEntity> serviceAreaList = servicePointChargeRptMapper.getServicePointServiceAreasNew(search);
        Map<Long, List<MDServicePointArea>> serviceAreaMap = serviceAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId));

        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "paymentType", "primaryId", "address", "remarks"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
        MDServicePointViewModel servicePointVM;

        List<Long> engineerIds = servicePointMap.values().stream().map(MDServicePointViewModel::getPrimaryId).distinct().collect(Collectors.toList());

        Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
        if (engineerIds != null && !engineerIds.isEmpty()) {
            List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Arrays.asList("id", "name", "appFlag"));
            if (engineerList != null && !engineerList.isEmpty()) {
                engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
            }
        }
        RPTEngineer engineer;
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        RPTServicePoint servicePoint;
        RPTDict engineerPaymentTypeDict;
        for (RPTServicePointPaySummaryEntity entity : returnPage) {
            servicePointVM = servicePointMap.get(entity.getServicePointId());
            if (servicePointVM != null) {
                servicePoint = new RPTServicePoint();
                servicePoint.setServicePointNo(servicePointVM.getServicePointNo());
                servicePoint.setName(servicePointVM.getName());
                servicePoint.setContactInfo1(servicePointVM.getContactInfo1());
                servicePoint.setContactInfo2(servicePointVM.getContactInfo2());

                if (servicePointVM.getPaymentType() != null) {
                    engineerPaymentTypeDict = paymentTypeMap.get(servicePointVM.getPaymentType().toString());
                    if (engineerPaymentTypeDict != null) {
                        servicePoint.setPaymentType(engineerPaymentTypeDict);
                    }
                }
                entity.setServicePoint(servicePoint);
                engineer = engineerMap.get(servicePointVM.getPrimaryId());
                if (engineer != null) {
                    entity.setPrimaryName(engineer.getName());
                    entity.setAppFlag(engineer.getAppFlag());
                }
                entity.setAddress(servicePointVM.getAddress());
                StringBuilder sb = new StringBuilder();
                if (serviceAreaMap.get(entity.getServicePointId()) != null) {
                    int index = 0;
                    for(MDServicePointArea item : serviceAreaMap.get(entity.getServicePointId())){
                        RPTArea  rptArea = cacheCountyAreaMap.get(item.getAreaId());
                        if(index != 0){
                            sb.append(",");
                        }
                        if(rptArea != null){
                            index++;
                            sb.append(rptArea.getName());
                        }

                    }
                    entity.setServiceAreas(sb.toString());
                }
                entity.setServicePointRemarks(servicePointVM.getRemarks() == null ? "" : servicePointVM.getRemarks());
            }
        }
        return returnPage;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasServicePointPaySummaryReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointPaySummarySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointPaySummarySearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        searchCondition.setYearMonth(yearMonth);
        searchCondition.setQuarter(quarter);
        Integer rowCount = servicePointChargeRptMapper.hasServicePointPaySummaryReportData(searchCondition);
        result = rowCount > 0;

        return result;
    }


    /**
     * 网点应付款的导出
     */
    public SXSSFWorkbook servicePointPaySummaryExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointPaySummarySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointPaySummarySearch.class);
            List<RPTServicePointPaySummaryEntity> list = getServicePointPaySummaryRptData(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            list = list.stream().sorted(Comparator.comparing(i -> i.getServicePoint() == null ? "" : i.getServicePoint().getServicePointNo() == null ? "" : i.getServicePoint().getServicePointNo())).collect(Collectors.toList());
            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 33));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 7));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务网点信息");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 8, 8));
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完成单");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 27));
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "费用情况");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 28, 32));
            ExportExcel.createCell(headFirstRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "银行账号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 33, 33));
            ExportExcel.createCell(headFirstRow, 33, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点店名");
            ExportExcel.createCell(headSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点负责人");
            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话1");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话2");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");

            ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月余额");
            ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月预付");
            ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(headSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效奖励");
            ExportExcel.createCell(headSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效费");
            ExportExcel.createCell(headSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");

            ExportExcel.createCell(headSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台服务费(入账)");
            ExportExcel.createCell(headSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月互助基金");
            ExportExcel.createCell(headSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退补金额");
            ExportExcel.createCell(headSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月应付合计");
            ExportExcel.createCell(headSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月实付");
            ExportExcel.createCell(headSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台服务费(付款)");
            ExportExcel.createCell(headSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月余额");
            ExportExcel.createCell(headSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "师傅质保金额");
            ExportExcel.createCell(headSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保金额");
            ExportExcel.createCell(headSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费用");
            ExportExcel.createCell(headSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");

            ExportExcel.createCell(headSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开票");
            ExportExcel.createCell(headSecondRow, 29, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");
            ExportExcel.createCell(headSecondRow, 30, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "所属银行及支行");
            ExportExcel.createCell(headSecondRow, 31, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "账户名");
            ExportExcel.createCell(headSecondRow, 32, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "账号");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {

                double totalPreBalance = 0.0;
                double totalPreDeposit = 0.0;
                double totalCompletedOrderCharge = 0.0;
                double totalTimelinessCharge = 0.0;
                double totalCustomerTimelinessCharge = 0.0;
                double totalUrgentCharge = 0.0;
                double praiseFee = 0.0;
                double infoFee = 0.0;
                double engineerDeposit = 0.0;
                double rechargeDeposit = 0.0;
                double taxFee = 0.0;
                double totalInsuranceCharge = 0.0;
                double totalDiffCharge = 0.0;
                double totalPayableAmount = 0.0;
                double totalPaidAmount = 0.0;
                double totalPlatformFee = 0.0;
                double totalTheBalance = 0.0;
                double totalTravelCharge = 0.0;
                double totalOtherCharge = 0.0;
                int totalFinishQty = 0;

                RPTServicePointPaySummaryEntity rowData = null;
                Row dataRow = null;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    rowData = list.get(dataRowIndex);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getServicePointNo());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getPrimaryName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getContactInfo1());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getContactInfo2());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getAddress());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getPaymentType() == null ? "" : rowData.getServicePoint().getPaymentType().getLabel());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompleteQty());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getLastMonthBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPreDeposit());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompletedCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTimelinessCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerTimelinessCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUrgentCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTaxFee());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPraiseFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getInfoFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getInsuranceCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReturnCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPayableAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getActualPaidAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPlatformFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTheBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerDeposit());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRechargeDeposit());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerTravelCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerOtherCharge());

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getInvoiceFlag() == 0 ? "否" : "是");
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getDiscountFlag() == 0 ? "否" : "是");
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getBank() == null ? "" : rowData.getServicePoint().getBank().getLabel());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getBankOwner());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getBankNo());

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex);
                    dataCell.setCellValue(rowData.getServicePointRemarks() == null ? "" : rowData.getServicePointRemarks());

                    totalPreBalance = totalPreBalance + rowData.getLastMonthBalance();
                    totalPreDeposit = totalPreDeposit + rowData.getPreDeposit();
                    totalCompletedOrderCharge = totalCompletedOrderCharge + rowData.getCompletedCharge();
                    totalTimelinessCharge = totalTimelinessCharge + rowData.getTimelinessCharge();
                    totalCustomerTimelinessCharge = totalCustomerTimelinessCharge + rowData.getCustomerTimelinessCharge();
                    totalUrgentCharge = totalUrgentCharge + rowData.getUrgentCharge();
                    praiseFee = praiseFee + rowData.getPraiseFee();
                    infoFee = infoFee + rowData.getInfoFee();
                    engineerDeposit = engineerDeposit + rowData.getEngineerDeposit();
                    rechargeDeposit = rechargeDeposit + rowData.getRechargeDeposit();
                    taxFee = taxFee + rowData.getTaxFee();
                    totalInsuranceCharge = totalInsuranceCharge + rowData.getInsuranceCharge();
                    totalDiffCharge = totalDiffCharge + rowData.getReturnCharge();
                    totalPayableAmount = totalPayableAmount + rowData.getPayableAmount();
                    totalPaidAmount = totalPaidAmount + rowData.getActualPaidAmount();
                    totalPlatformFee = totalPlatformFee + rowData.getPlatformFee();
                    totalTheBalance = totalTheBalance + rowData.getTheBalance();
                    totalTravelCharge = totalTravelCharge + rowData.getEngineerTravelCharge();
                    totalOtherCharge = totalOtherCharge + rowData.getEngineerOtherCharge();
                    totalFinishQty = totalFinishQty + rowData.getCompleteQty();
                }


                Row sumRow = xSheet.createRow(rowIndex);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 7));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalFinishQty);
                ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPreBalance);
                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPreDeposit);
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCompletedOrderCharge);
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTimelinessCharge);
                ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCustomerTimelinessCharge);
                ExportExcel.createCell(sumRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalUrgentCharge);
                ExportExcel.createCell(sumRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, taxFee);

                ExportExcel.createCell(sumRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
                ExportExcel.createCell(sumRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, infoFee);
                ExportExcel.createCell(sumRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalInsuranceCharge);
                ExportExcel.createCell(sumRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalDiffCharge);
                ExportExcel.createCell(sumRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPayableAmount);
                ExportExcel.createCell(sumRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPaidAmount);
                ExportExcel.createCell(sumRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPlatformFee);
                ExportExcel.createCell(sumRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTheBalance);
                ExportExcel.createCell(sumRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);
                ExportExcel.createCell(sumRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rechargeDeposit);
                ExportExcel.createCell(sumRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTravelCharge);
                ExportExcel.createCell(sumRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOtherCharge);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 28, 33));
                ExportExcel.createCell(sumRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            }

        } catch (Exception e) {
            log.error("【ServicePointChargeRptService.servicePointPaySummaryExport】网点应付汇总报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasServicePointCostPerReportData(String searchConditionJson) {
        boolean result = true;
        RPTServicePointPaySummarySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointPaySummarySearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        searchCondition.setYearMonth(yearMonth);
        searchCondition.setQuarter(quarter);
        Integer rowCount = servicePointChargeRptMapper.hasServicePointCostPerReportData(searchCondition);
        result = rowCount > 0;

        return result;
    }

    /**
     * 网点成本排名的导出
     */
    public SXSSFWorkbook servicePointCostPerExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointPaySummarySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointPaySummarySearch.class);
            List<RPTServicePointPaySummaryEntity> list = getServicePointCostPerRptList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 28));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 9));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务网点信息");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 10, 10));
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完成单");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 11, 26));
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "费用情况");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 27, 27));
            ExportExcel.createCell(headFirstRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平均每单成本");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 28, 28));
            ExportExcel.createCell(headFirstRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点店名");
            ExportExcel.createCell(headSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点负责人");
            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话1");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话2");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "自行接单");
            ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "负责区域");

            ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月余额");
            ExportExcel.createCell(headSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(headSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效奖励");
            ExportExcel.createCell(headSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效费");
            ExportExcel.createCell(headSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");
            ExportExcel.createCell(headSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "扣点");
            ExportExcel.createCell(headSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平台费");
            ExportExcel.createCell(headSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "质保金额");
            ExportExcel.createCell(headSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
            ExportExcel.createCell(headSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月互助基金");
            ExportExcel.createCell(headSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退补金额");
            ExportExcel.createCell(headSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月应付合计");
            ExportExcel.createCell(headSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月实付");
            ExportExcel.createCell(headSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月余额");
            ExportExcel.createCell(headSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费用");
            ExportExcel.createCell(headSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {

                double totalPreBalance = 0.0;
                double totalCompletedOrderCharge = 0.0;
                double totalTimelinessCharge = 0.0;
                double totalCustomerTimelinessCharge = 0.0;
                double totalUrgentCharge = 0.0;
                double praiseFee = 0.0;
                double taxFee = 0.0;
                double infoFee = 0.0;
                double engineerDeposit = 0.0;
                double totalInsuranceCharge = 0.0;
                double totalDiffCharge = 0.0;
                double totalPayableAmount = 0.0;
                double totalPaidAmount = 0.0;
                double totalTheBalance = 0.0;
                double totalTravelCharge = 0.0;
                double totalOtherCharge = 0.0;
                int totalFinishQty = 0;

                RPTServicePointPaySummaryEntity rowData = null;
                Row dataRow = null;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    rowData = list.get(dataRowIndex);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);

                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getServicePointNo());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getPrimaryName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getContactInfo1());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getContactInfo2());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getAddress());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePoint() == null ? "" : rowData.getServicePoint().getPaymentType() == null ? "" : rowData.getServicePoint().getPaymentType().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAppFlag() == 1 ? "是" : "否");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServiceAreas());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompleteQty());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getLastMonthBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompletedCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTimelinessCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerTimelinessCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUrgentCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTaxFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getInfoFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerDeposit());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPraiseFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getInsuranceCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReturnCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPayableAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPaidAmount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTheBalance());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerTravelCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerOtherCharge());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCostPerOrder());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointRemarks() == null ? "" : rowData.getServicePointRemarks());

                    totalPreBalance = totalPreBalance + rowData.getLastMonthBalance();
                    totalCompletedOrderCharge = totalCompletedOrderCharge + rowData.getCompletedCharge();
                    totalTimelinessCharge = totalTimelinessCharge + rowData.getTimelinessCharge();
                    totalCustomerTimelinessCharge = totalCustomerTimelinessCharge + rowData.getCustomerTimelinessCharge();
                    totalUrgentCharge = totalUrgentCharge + rowData.getUrgentCharge();
                    praiseFee = praiseFee + rowData.getPraiseFee();
                    taxFee = taxFee + rowData.getTaxFee();
                    infoFee = infoFee + rowData.getInfoFee();
                    engineerDeposit = engineerDeposit + rowData.getEngineerDeposit();
                    totalInsuranceCharge = totalInsuranceCharge + rowData.getInsuranceCharge();
                    totalDiffCharge = totalDiffCharge + rowData.getReturnCharge();
                    totalPayableAmount = totalPayableAmount + rowData.getPayableAmount();
                    totalPaidAmount = totalPaidAmount + rowData.getPaidAmount();
                    totalTheBalance = totalTheBalance + rowData.getTheBalance();
                    totalTravelCharge = totalTravelCharge + rowData.getEngineerTravelCharge();
                    totalOtherCharge = totalOtherCharge + rowData.getEngineerOtherCharge();
                    totalFinishQty = totalFinishQty + rowData.getCompleteQty();
                }

                Row sumRow = xSheet.createRow(rowIndex);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 9));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalFinishQty);
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPreBalance);
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCompletedOrderCharge);
                ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTimelinessCharge);
                ExportExcel.createCell(sumRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCustomerTimelinessCharge);
                ExportExcel.createCell(sumRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalUrgentCharge);
                ExportExcel.createCell(sumRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, taxFee);
                ExportExcel.createCell(sumRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, infoFee);
                ExportExcel.createCell(sumRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerDeposit);
                ExportExcel.createCell(sumRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, praiseFee);
                ExportExcel.createCell(sumRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalInsuranceCharge);
                ExportExcel.createCell(sumRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalDiffCharge);
                ExportExcel.createCell(sumRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPayableAmount);
                ExportExcel.createCell(sumRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPaidAmount);
                ExportExcel.createCell(sumRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTheBalance);
                ExportExcel.createCell(sumRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTravelCharge);
                ExportExcel.createCell(sumRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOtherCharge);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 27, 28));
                ExportExcel.createCell(sumRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            }

        } catch (Exception e) {
            log.error("【ServicePointChargeRptService.servicePointCostPerExport】网点成本排名报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 写入中间表
     *
     * @param date
     * @return
     */
    private List<RPTServicePointChargeEntity> getServicePointChargeData(Date date) {
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        if (endDate.getTime() > new Date().getTime()) {
            endDate = DateUtils.getDateStart(new Date());
        }
        int selectedYear = DateUtils.getYear(date);
        int selectedMonth = DateUtils.getMonth(date);
        Date preMonthQueryDate = DateUtils.addMonth(date, -1);
        List<RPTServicePointChargeEntity> list = new ArrayList<>();
        int endLimit = 20000;
        int beginLimit;
        int pageNo;
        int listSize = 0;
        List<RPTServicePointChargeEntity> payableAAllList = Lists.newArrayList();
        List<RPTServicePointChargeEntity> payableBAllList = Lists.newArrayList();
        List<RPTServicePointChargeEntity> payableAList;
        List<RPTServicePointChargeEntity> payableBList;

        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            payableAList = servicePointChargeRptMapper.getPayableChargeA(beginDate, endDate, quarter, beginLimit, endLimit);
            if (payableAList != null) {
                payableAAllList.addAll(payableAList);
                listSize = payableAList.size();
            }

            pageNo++;
        } while (listSize == endLimit);

        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            payableBList = servicePointChargeRptMapper.getPayableChargeB(beginDate, endDate, quarter, beginLimit, endLimit);
            if (payableBList != null) {
                payableBAllList.addAll(payableBList);
                listSize = payableBList.size();
            }

            pageNo++;
        } while (listSize == endLimit);

        Map<String, RPTServicePointChargeEntity> payableAMap = payableAAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        Map<String, RPTServicePointChargeEntity> payableBMap = payableBAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> paidAllList = Lists.newArrayList();
        List<RPTServicePointChargeEntity> paidList;
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            paidList = servicePointChargeRptMapper.getServicePointTotalPaid(selectedYear, selectedMonth, beginLimit, endLimit);
            if (paidList != null) {
                paidAllList.addAll(paidList);
                listSize = paidList.size();
            }

            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> paidMap = paidAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> currBalanceList;
        List<RPTServicePointChargeEntity> currBalanceAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            currBalanceList = servicePointChargeRptMapper.getServicePointTotalCurrBalance(selectedYear, selectedMonth, beginLimit, endLimit);
            if (currBalanceList != null) {
                currBalanceAllList.addAll(currBalanceList);
                listSize = currBalanceList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> currBalanceMap = currBalanceAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> completedChargeList;
        List<RPTServicePointChargeEntity> completedChargeAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            completedChargeList = servicePointChargeRptMapper.getServicePointCompletedOrderCharge(beginDate, endDate, quarter, beginLimit, endLimit);
            if (completedChargeList != null) {
                completedChargeAllList.addAll(completedChargeList);
                listSize = completedChargeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> completedChargeMap = completedChargeAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> timelinessAndInsuranceChargeList;
        List<RPTServicePointChargeEntity> timelinessAndInsuranceChargeAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            timelinessAndInsuranceChargeList = servicePointChargeRptMapper.getServicePointTimelinessAndInsurance(beginDate, endDate, quarter, beginLimit, endLimit);
            if (timelinessAndInsuranceChargeList != null) {
                timelinessAndInsuranceChargeAllList.addAll(timelinessAndInsuranceChargeList);
                listSize = timelinessAndInsuranceChargeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> timelinessAndInsuranceChargeMap = timelinessAndInsuranceChargeAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> diffChargeList;
        List<RPTServicePointChargeEntity> diffChargeAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            diffChargeList = servicePointChargeRptMapper.getServicePointDiffCharge(beginDate, endDate, quarter, beginLimit, endLimit);
            if (diffChargeList != null) {
                diffChargeAllList.addAll(diffChargeList);
                listSize = diffChargeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> diffChargeMap = diffChargeAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> preBalanceList;
        List<RPTServicePointChargeEntity> preBalanceAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            preBalanceList = servicePointChargeRptMapper.getServicePointTotalBalance(DateUtils.getYear(preMonthQueryDate), DateUtils.getMonth(preMonthQueryDate), beginLimit, endLimit);
            if (preBalanceList != null) {
                preBalanceAllList.addAll(preBalanceList);
                listSize = preBalanceList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> preBalanceMap = preBalanceAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> travelAndOtherChargeList;
        List<RPTServicePointChargeEntity> travelAndOtherChargeAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            travelAndOtherChargeList = servicePointChargeRptMapper.getServicePointOtherChargeAndTravelCharge(beginDate, endDate, quarter, beginLimit, endLimit);
            if (travelAndOtherChargeList != null) {
                travelAndOtherChargeAllList.addAll(travelAndOtherChargeList);
                listSize = travelAndOtherChargeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> travelAndOtherChargeMap = travelAndOtherChargeAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));

        List<RPTServicePointChargeEntity> completeQtyList;
        List<RPTServicePointChargeEntity> completeQtyAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            completeQtyList = servicePointChargeRptMapper.getAllServicePointFinishQty(beginDate, endDate, beginLimit, endLimit);
            if (completeQtyList != null) {
                completeQtyAllList.addAll(completeQtyList);
                listSize = completeQtyList.size();
            }
            pageNo++;
        } while (listSize == endLimit);
        Map<String, RPTServicePointChargeEntity> completeQtyMap = completeQtyAllList.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId(), Function.identity()));
        List<RPTServicePointChargeEntity> platformFeeList;
        List<RPTServicePointChargeEntity> platformFeeAllList = Lists.newArrayList();
        pageNo = 1;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            platformFeeList = servicePointChargeRptMapper.getAllServicePointPlatFees(quarter, beginDate, endDate, beginLimit, endLimit);
            if (platformFeeList != null) {
                platformFeeAllList.addAll(platformFeeList);
                listSize = platformFeeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);

        for (RPTServicePointChargeEntity pointChargeEntity : platformFeeAllList) {
            pointChargeEntity.setProductCategoryId(0L);
        }
        Map<Long, RPTServicePointChargeEntity> platformFeeMap = platformFeeAllList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        Set<String> keySet = Sets.newHashSet();
        keySet.addAll(payableAMap.keySet());
        keySet.addAll(payableBMap.keySet());
        keySet.addAll(paidMap.keySet());
        keySet.addAll(currBalanceMap.keySet());
        keySet.addAll(completedChargeMap.keySet());
        keySet.addAll(timelinessAndInsuranceChargeMap.keySet());
        keySet.addAll(diffChargeMap.keySet());
        keySet.addAll(preBalanceMap.keySet());
        keySet.addAll(travelAndOtherChargeMap.keySet());
        keySet.addAll(completeQtyMap.keySet());


        Set<Long> servicePointIds = Sets.newHashSet();
        servicePointIds.addAll(payableAAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(payableBAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(paidAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(currBalanceAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(completedChargeAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(timelinessAndInsuranceChargeAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(diffChargeAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(preBalanceAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(travelAndOtherChargeAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(completeQtyAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));
        servicePointIds.addAll(platformFeeAllList.stream().map(RPTServicePointChargeEntity::getServicePointId).collect(Collectors.toSet()));

        List<Long> servicePointIdList = Lists.newArrayList(servicePointIds);
        String[] fieldsArray = new String[]{"id", "primaryId", "paymentType", "areaId"};
        Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIdList, Arrays.asList(fieldsArray), null);
        Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
        if (servicePointMap != null) {
            List<Long> engineerIds = servicePointMap.values().stream().map(MDServicePointViewModel::getPrimaryId).distinct().collect(Collectors.toList());
            if (engineerIds != null && !engineerIds.isEmpty()) {
                List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Arrays.asList("id", "name", "appFlag"));
                if (engineerList != null && !engineerList.isEmpty()) {
                    engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
                }
            }
        }
        List<String> keyList = new ArrayList<>(keySet);
        RPTServicePointChargeEntity entity;
        RPTServicePointChargeEntity chargeEntity;
        Long productCategoryId;
        Long servicePointId;
        double platFee;
        RPTEngineer engineer;
        MDServicePointViewModel servicePoint;
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
        for (String key : keyList) {
            entity = new RPTServicePointChargeEntity();
            entity.setSystemId(systemId);
            entity.setYearMonth(yearMonth);
            entity.setQuarter(quarter);
            chargeEntity = payableAMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setPayableA(chargeEntity.getPayableA());
            }
            chargeEntity = payableBMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setPayableB(chargeEntity.getPayableB());
            }
            chargeEntity = paidMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setPaidAmount(chargeEntity.getPaidAmount());
                //设置平台服务费
                platFee = CurrencyUtil.round2(entity.getPaidAmount() * RptCommonUtils.getPlatformFeeRate());
                entity.setPlatformFee(platFee);


            }
            chargeEntity = currBalanceMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setTheBalance(chargeEntity.getTheBalance());

            }
            chargeEntity = completedChargeMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setCompletedCharge(chargeEntity.getCompletedCharge());
            }
            chargeEntity = timelinessAndInsuranceChargeMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setTimelinessCharge(chargeEntity.getTimelinessCharge());
                entity.setCustomerTimelinessCharge(chargeEntity.getCustomerTimelinessCharge());
                entity.setUrgentCharge(chargeEntity.getUrgentCharge());
                entity.setInsuranceCharge(chargeEntity.getInsuranceCharge());
                entity.setPraiseFee(chargeEntity.getPraiseFee());
                entity.setTaxFee(chargeEntity.getTaxFee());
                entity.setInfoFee(chargeEntity.getInfoFee());
                entity.setEngineerDeposit(chargeEntity.getEngineerDeposit());
            }
            chargeEntity = diffChargeMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setReturnCharge(chargeEntity.getReturnCharge());
            }
            chargeEntity = preBalanceMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setLastMonthBalance(chargeEntity.getLastMonthBalance());
            }
            chargeEntity = travelAndOtherChargeMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setEngineerOtherCharge(chargeEntity.getEngineerOtherCharge());
                entity.setEngineerTravelCharge(chargeEntity.getEngineerTravelCharge());

            }
            chargeEntity = completeQtyMap.get(key);
            if (chargeEntity != null) {
                productCategoryId = chargeEntity.getProductCategoryId();
                servicePointId = chargeEntity.getServicePointId();
                entity.setServicePointId(servicePointId);
                entity.setProductCategoryId(productCategoryId);
                entity.setCompleteQty(chargeEntity.getCompleteQty());
            }
            entity.setPayableAmount(entity.getPayableA() + entity.getPayableB());
            entity.getCostPerOrder();
            list.add(entity);
        }
        RPTArea area;
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, List<RPTServicePointChargeEntity>> allTravelMap = list.stream().collect(Collectors.groupingBy(RPTServicePointChargeEntity::getServicePointId));
        RPTServicePointChargeEntity item;

        for (RPTServicePointChargeEntity rptEntity : list) {
            if (servicePointMap != null) {
                servicePoint = servicePointMap.get(rptEntity.getServicePointId());
                if (servicePoint != null) {
                    area = areaMap.get(servicePoint.getAreaId());
                    rptEntity.setPaymentType(servicePoint.getPaymentType());
                    if (area != null) {
                        String[] split = area.getParentIds().split(",");
                        if (split.length == 4) {
                            rptEntity.setCountyId(area.getId());
                            rptEntity.setCityId(Long.valueOf(split[3]));
                            rptEntity.setProvinceId(Long.valueOf(split[2]));
                        }
                    }
                    engineer = engineerMap.get(servicePoint.getPrimaryId());
                    if (engineer != null) {
                        rptEntity.setAppFlag(engineer.getAppFlag());
                    }
                }

            }

        }


        int completeQty;
        double lastMonthBalance;
        double preDeposit;
        double completedCharge;
        double timelinessCharge;
        double customerTimelinessCharge;
        double urgentCharge;
        double praiseFee;
        double infoFee;
        double engineerDeposit;
        double taxFee;
        double insuranceCharge;
        double returnCharge;
        double payableAmount;
        double paidAmount;
        double platformTotalFee;
        double theBalance;
        double engineerTravelCharge;
        double engineerOtherCharge;
        double costPerOrder;
        RPTServicePointChargeEntity servicePointChargeEntity;
        List<RPTServicePointChargeEntity> allTravelChargeList = Lists.newArrayList();
        for (List<RPTServicePointChargeEntity> entityList : allTravelMap.values()) {
            item = new RPTServicePointChargeEntity();
            servicePointChargeEntity = entityList.get(0);
            item.setProductCategoryId(0L);
            item.setServicePointId(servicePointChargeEntity.getServicePointId());
            item.setPaymentType(servicePointChargeEntity.getPaymentType());
            item.setProvinceId(servicePointChargeEntity.getProvinceId());
            item.setCityId(servicePointChargeEntity.getCityId());
            item.setCountyId(servicePointChargeEntity.getCountyId());
            item.setYearMonth(servicePointChargeEntity.getYearMonth());
            item.setSystemId(servicePointChargeEntity.getSystemId());
            item.setQuarter(servicePointChargeEntity.getQuarter());
            item.setAppFlag(servicePointChargeEntity.getAppFlag());
            completeQty = 0;
            lastMonthBalance = 0.0;
            preDeposit = 0.0;
            completedCharge = 0.0;
            timelinessCharge = 0.0;
            customerTimelinessCharge = 0.0;
            urgentCharge = 0.0;
            praiseFee = 0.0;
            infoFee = 0.0;
            taxFee = 0.0;
            engineerDeposit = 0.0;
            insuranceCharge = 0.0;
            returnCharge = 0.0;
            payableAmount = 0.0;
            paidAmount = 0.0;
            platformTotalFee = 0.0;
            theBalance = 0.0;
            engineerTravelCharge = 0.0;
            engineerOtherCharge = 0.0;
            costPerOrder = 0.0;
            if (platformFeeMap.get(item.getServicePointId()) != null) {
                platformTotalFee = platformFeeMap.get(item.getServicePointId()).getPlatformFee();
            }
            item.setPlatformFee(0 - platformTotalFee);

            for (int i = 0; i < entityList.size(); i++) {
                RPTServicePointChargeEntity pointChargeEntity = entityList.get(i);

                completeQty += pointChargeEntity.getCompleteQty();
                lastMonthBalance += pointChargeEntity.getLastMonthBalance();
                preDeposit += pointChargeEntity.getPreDeposit();
                completedCharge += pointChargeEntity.getCompletedCharge();
                timelinessCharge += pointChargeEntity.getTimelinessCharge();
                customerTimelinessCharge += pointChargeEntity.getCustomerTimelinessCharge();
                urgentCharge += pointChargeEntity.getUrgentCharge();
                praiseFee += pointChargeEntity.getPraiseFee();
                infoFee += pointChargeEntity.getInfoFee();
                engineerDeposit += pointChargeEntity.getEngineerDeposit();
                taxFee += pointChargeEntity.getTaxFee();
                insuranceCharge += pointChargeEntity.getInsuranceCharge();
                returnCharge += pointChargeEntity.getReturnCharge();
                payableAmount += pointChargeEntity.getPayableAmount();
                paidAmount += pointChargeEntity.getPaidAmount();
                theBalance += pointChargeEntity.getTheBalance();
                engineerTravelCharge += pointChargeEntity.getEngineerTravelCharge();
                engineerOtherCharge += pointChargeEntity.getEngineerOtherCharge();
                if (i == entityList.size() - 1) {
                    pointChargeEntity.setPlatformFee(0 - platformTotalFee);
                }
                platformTotalFee = platformTotalFee + pointChargeEntity.getPlatformFee();


            }


            item.setCompleteQty(completeQty);
            item.setLastMonthBalance(lastMonthBalance);
            item.setPreDeposit(preDeposit);
            item.setCompletedCharge(completedCharge);
            item.setTimelinessCharge(timelinessCharge);
            item.setCustomerTimelinessCharge(customerTimelinessCharge);
            item.setUrgentCharge(urgentCharge);
            item.setPraiseFee(praiseFee);
            item.setInfoFee(infoFee);
            item.setEngineerDeposit(engineerDeposit);
            item.setTaxFee(taxFee);
            item.setInsuranceCharge(insuranceCharge);
            item.setReturnCharge(returnCharge);
            item.setPayableAmount(payableAmount);
            item.setPaidAmount(paidAmount);
            item.setTheBalance(theBalance);
            item.setEngineerTravelCharge(engineerTravelCharge);
            item.setEngineerOtherCharge(engineerOtherCharge);
            if (completeQty != 0) {
                costPerOrder = payableAmount / completeQty;
            }
            item.setCostPerOrder(costPerOrder);

            allTravelChargeList.add(item);

        }
        list.addAll(allTravelChargeList);

        return list;
    }

    public void saveServicePointChargeToRptDB(Date date) {
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();

        boolean InfoFeeEnabled = RptCommonUtils.getInfoFeeEnabled();
        String quarter = QuarterUtils.getSeasonQuarter(date);

        List<RPTServicePointChargeEntity> tuple = servicePointChargeRptMapper.getUpServicePointChargeData(systemId, yearMonth, quarter);
        Map<String, RPTServicePointChargeEntity> tupleMaps = tuple.stream().collect(Collectors.toMap(i -> i.getServicePointId() + ":" + i.getProductCategoryId() + ":" + i.getSystemId() + ":" + i.getYearMonth(), Function.identity()));

        List<RPTServicePointChargeEntity> list = getServicePointChargeData(date);
        String key;
        RPTServicePointChargeEntity upEntity;
        if (!list.isEmpty()) {
            try {
                for (RPTServicePointChargeEntity entity : list) {
                    key = entity.getServicePointId() + ":" + entity.getProductCategoryId() + ":" + systemId + ":" + entity.getYearMonth();
                    upEntity = tupleMaps.get(key);
                    if (upEntity != null) {
                        entity.setId(upEntity.getId());
                        if (InfoFeeEnabled) {
                            servicePointChargeRptMapper.updateServicePointChargeInfoFee(entity);
                        } else {
                            servicePointChargeRptMapper.updateServicePointCharge(entity);
                        }

                    } else {
                        servicePointChargeRptMapper.insertServicePointCharge(entity);
                    }

                }
            } catch (Exception e) {
                log.error("【ServicePointChargeRptService.saveServicePointChargeToRptDB】网点写入中间表失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            }
        }
    }

    /**
     * 删除中间表中指定日期的客户每日催单数据
     */
    private void deleteServicePointChargeRptDB(Date date) {
        if (date != null) {
            int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
            int systemId = RptCommonUtils.getSystemId();
            String quarter = QuarterUtils.getSeasonQuarter(date);
            servicePointChargeRptMapper.deleteServicePointCharge(systemId, yearMonth, quarter);
        }
    }


    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() <= endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveServicePointChargeToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteServicePointChargeRptDB(beginDate);
                            saveServicePointChargeToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteServicePointChargeRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointChargeRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 重建中间表
     */
    public boolean rebuildUpdatePlatFromFee(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() <= endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            updatePlatFromFee(beginDate);
                            break;
                        case DELETE:
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointChargeRptService.rebuildUpdatePlatFromFee:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    public void updatePlatFromFee(Date date) {
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        if (endDate.getTime() > new Date().getTime()) {
            endDate = DateUtils.getDateStart(new Date());
        }
        List<RPTServicePointChargeEntity> list = servicePointChargeRptMapper.getUpdatePlatformFeeData(systemId, yearMonth, quarter);


        List<RPTServicePointChargeEntity> platformFeeList;
        List<RPTServicePointChargeEntity> platformFeeAllList = Lists.newArrayList();
        int endLimit = 20000;
        int beginLimit;
        int pageNo = 1;
        int listSize = 0;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            platformFeeList = servicePointChargeRptMapper.getAllServicePointPlatFees(quarter, beginDate, endDate, beginLimit, endLimit);
            if (platformFeeList != null) {
                platformFeeAllList.addAll(platformFeeList);
                listSize = platformFeeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);

        Map<Long, List<RPTServicePointChargeEntity>> allTravelMap = list.stream().collect(Collectors.groupingBy(RPTServicePointChargeEntity::getServicePointId));

        Map<Long, RPTServicePointChargeEntity> platformFeeMap = platformFeeAllList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));
        if (platformFeeMap == null) {
            return;
        }
        double platFee;
        double platformTotalFee;

        for (List<RPTServicePointChargeEntity> entityList : allTravelMap.values()) {

            platformTotalFee = 0;

            if (platformFeeMap.get(entityList.get(0).getServicePointId()) != null) {
                platformTotalFee = platformFeeMap.get(entityList.get(0).getServicePointId()).getPlatformFee();
            }

            for (int i = 0; i < entityList.size(); i++) {
                RPTServicePointChargeEntity entity = entityList.get(i);
                if (platformFeeMap.get(entity.getServicePointId()) != null) {

                    //设置平台服务费
                    platFee = CurrencyUtil.round2(entity.getPaidAmount() * RptCommonUtils.getPlatformFeeRate());
                    entity.setPlatformFee(platFee);

                    if (i == entityList.size() - 1) {
                        entity.setPlatformFee(0 - platformTotalFee);
                    }

                    platformTotalFee = platformTotalFee + platFee;


                    servicePointChargeRptMapper.updatePlatformFee(entity);
                }

            }

        }
    }


    /**
     * 重建中间表
     */
    public boolean rebuildUpdateQuarterMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() <= endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            updateServicePointChartQuarter(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            updateServicePointChartQuarter(beginDate);
                            break;
                        case DELETE:
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }

                result = true;
            } catch (Exception e) {
                log.error("ServicePointChargeRptService.rebuildUpdateQuarterMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    public void updateServicePointChartQuarter(Date date) {
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        if (quarter.equals("20171")) {
            return;
        }
        List<RPTServicePointChargeEntity> list = servicePointChargeRptMapper.getUpdateServicePointQuarter(systemId, yearMonth, quarter);

        for (RPTServicePointChargeEntity entity : list) {

            servicePointChargeRptMapper.insertServicePointChargeNew(entity);
        }
    }

    /**
     * 重建中间表
     */
    public boolean rebuildUpdatePlatFromFeeNew(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = DateUtils.getStartDayOfMonth(new Date(beginDt));
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() <= endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            updatePlatFromFeeNew(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            updatePlatFromFeeNew(beginDate);
                            break;
                        case DELETE:
                            break;
                    }
                    beginDate = DateUtils.addMonth(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("ServicePointChargeRptService.rebuildUpdatePlatFromFeeNew:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    public void updatePlatFromFeeNew(Date date) {
        int yearMonth = StringUtils.toInteger(DateUtils.formatDate(date, "yyyyMM"));
        int systemId = RptCommonUtils.getSystemId();
        String quarter = QuarterUtils.getSeasonQuarter(date);
        Date beginDate = DateUtils.getStartDayOfMonth(date);
        Date endDate = DateUtils.getLastDayOfMonth(date);
        if (endDate.getTime() > new Date().getTime()) {
            endDate = DateUtils.getDateStart(new Date());
        }

        List<RPTServicePointChargeEntity> updatePlatformFeeList = servicePointChargeRptMapper.getUpdatePlatformFeeList(systemId, yearMonth, quarter);
        for (RPTServicePointChargeEntity entity : updatePlatformFeeList) {
            entity.setPlatformFee(0.0);
            servicePointChargeRptMapper.updatePlatformFee(entity);
        }

        Map<Long, List<RPTServicePointChargeEntity>> updatePlatformFeeMap = updatePlatformFeeList.stream().collect(Collectors.groupingBy(RPTServicePointChargeEntity::getServicePointId));
        List<RPTServicePointChargeEntity> platformFeeList;
        List<RPTServicePointChargeEntity> platformFeeAllList = Lists.newArrayList();
        int endLimit = 20000;
        int beginLimit;
        int pageNo = 1;
        int listSize = 0;
        do {
            beginLimit = (pageNo - 1) * endLimit;
            platformFeeList = servicePointChargeRptMapper.getAllServicePointPlatFees(quarter, beginDate, endDate, beginLimit, endLimit);
            if (platformFeeList != null) {
                platformFeeAllList.addAll(platformFeeList);
                listSize = platformFeeList.size();
            }
            pageNo++;
        } while (listSize == endLimit);

        Map<Long, RPTServicePointChargeEntity> platformFeeMap = platformFeeAllList.stream().collect(Collectors.toMap(RPTServicePointChargeEntity::getServicePointId, Function.identity()));

        double platformFee;

        for (Map.Entry<Long, List<RPTServicePointChargeEntity>> entityMap : updatePlatformFeeMap.entrySet()) {
            platformFee = 0;
            if (platformFeeMap != null) {
                if (platformFeeMap.get(entityMap.getKey()) != null) {
                    platformFee = platformFeeMap.get(entityMap.getKey()).getPlatformFee();
                }
            }

            if (platformFee != 0) {
                for (int i = 0; i < 2; i++) {
                    RPTServicePointChargeEntity entity = entityMap.getValue().get(i);
                    entity.setPlatformFee(0 - platformFee);
                    servicePointChargeRptMapper.updatePlatformFee(entity);
                }
            }

        }

    }
}
