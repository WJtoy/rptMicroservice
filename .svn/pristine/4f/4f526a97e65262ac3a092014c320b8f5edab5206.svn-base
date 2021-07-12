package com.kkl.kklplus.provider.rpt.customer.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.praise.PraiseStatusEnum;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.CustomerPraiseDetailsRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCustomerPraiseDetailsRptService extends RptBaseService {

    @Autowired
    private CustomerPraiseDetailsRptMapper customerPraiseDetailsRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private AreaCacheService areaCacheService;

    public Page<RPTKeFuPraiseDetailsEntity> getCustomerPraiseDetailsList(RPTKeFuCompleteTimeSearch search) {
        Page<RPTKeFuPraiseDetailsEntity> page;

        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);

        page = customerPraiseDetailsRptMapper.getCustomerPraiseDetailsPage(search);


        List<Long> customerIds = page.stream().map(RPTKeFuPraiseDetailsEntity::getCustomerId).distinct().collect(Collectors.toList());
        String[] customerFieldsArray = new String[]{"id", "name","salesId"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(customerFieldsArray));
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        Map<Long, String> namesByUserIds = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));

        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        String countyName = "";
        String cityName = "";
        String provinceName = "";
        RPTDict praiseDict;
        for(RPTKeFuPraiseDetailsEntity item : page){
            item.setCreateDate(new Date(item.getCreateDt()));

            RPTCustomer customer = customerMap.get(item.getCustomerId());
            if(customer != null) {
                item.setCustomerName(customer.getName());
                if (namesByUserIds != null && namesByUserIds.size() > 0) {
                    String salesName = namesByUserIds.get(customer.getSales().getId());
                    if (StringUtils.isNotBlank(salesName)) {
                        item.setKeFuName(salesName);
                    }
                }
            }


            if(item.getStatus() != null && item.getStatus().getValue() != null){
                praiseDict = new RPTDict();
                praiseDict.setLabel(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).msg);
                praiseDict.setValue(String.valueOf(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).code));
                item.setStatus(praiseDict);
            }
            RPTArea rptArea  = areaMap.get(item.getAreaId());
            if (rptArea != null) {
                countyName = rptArea.getName();
                RPTArea city = rptArea.getParent();
                if (city!=null){
                    RPTArea cityArea = cityMap.get(city.getId());
                    if (cityArea!=null){
                        RPTArea parent = cityArea.getParent();
                        cityName = cityArea.getName();
                        if (parent!=null){
                            RPTArea provinceArea = provinceMap.get(parent.getId());
                            if(provinceArea != null){
                                provinceName = provinceArea.getName();
                            }
                        }
                    }
                }
            }

            item.setAreaName(provinceName+cityName+countyName);
        }
        return page;
    }


    public List<RPTKeFuPraiseDetailsEntity> getCustomerPraiseDetail(RPTKeFuCompleteTimeSearch search) {
        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);

        List<RPTKeFuPraiseDetailsEntity> list = customerPraiseDetailsRptMapper.getCustomerPraiseDetailsPage(search);

        List<Long> customerIds = list.stream().map(RPTKeFuPraiseDetailsEntity::getCustomerId).distinct().collect(Collectors.toList());
        String[] customerFieldsArray = new String[]{"id", "name","salesId"};
        Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(customerFieldsArray));
        Set<Long> salesIds = customerMap.values().stream().filter(i -> i.getSales() != null && i.getSales().getId() != null)
                .map(i -> i.getSales().getId()).collect(Collectors.toSet());
        Map<Long, String> namesByUserIds = MSUserUtils.getNamesByUserIds(Lists.newArrayList(salesIds));

        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        String countyName = "";
        String cityName = "";
        String provinceName = "";
        RPTDict praiseDict;
        for(RPTKeFuPraiseDetailsEntity item : list){
            item.setCreateDate(new Date(item.getCreateDt()));

            RPTCustomer customer = customerMap.get(item.getCustomerId());
            if(customer != null) {
                item.setCustomerName(customer.getName());
                if (namesByUserIds != null && namesByUserIds.size() > 0) {
                    String salesName = namesByUserIds.get(customer.getSales().getId());
                    if (StringUtils.isNotBlank(salesName)) {
                        item.setKeFuName(salesName);
                    }
                }
            }

            if(item.getStatus() != null && item.getStatus().getValue() != null){
                praiseDict = new RPTDict();
                praiseDict.setLabel(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).msg);
                praiseDict.setValue(String.valueOf(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).code));
                item.setStatus(praiseDict);
            }
            RPTArea rptArea  = areaMap.get(item.getAreaId());
            if (rptArea != null) {
                countyName = rptArea.getName();
                RPTArea city = rptArea.getParent();
                if (city!=null){
                    RPTArea cityArea = cityMap.get(city.getId());
                    if (cityArea!=null){
                        RPTArea parent = cityArea.getParent();
                        cityName = cityArea.getName();
                        if (parent!=null){
                            RPTArea provinceArea = provinceMap.get(parent.getId());
                            if(provinceArea != null){
                                provinceName = provinceArea.getName();
                            }
                        }
                    }
                }
            }

            item.setAreaName(provinceName+cityName+countyName);
        }
        return list;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTKeFuCompleteTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuCompleteTimeSearch.class);
        if (searchCondition.getBeginDate() != null && searchCondition.getEndDate() != null) {
            Integer customerPraiseSum = customerPraiseDetailsRptMapper.getCustomerPraiseSum(searchCondition);
            result = customerPraiseSum > 0;
        }
        return result;
    }


    public SXSSFWorkbook customerPraiseDetailsRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTKeFuCompleteTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuCompleteTimeSearch.class);
            List<RPTKeFuPraiseDetailsEntity> list = getCustomerPraiseDetail(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单单号");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "销售单号");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务单号");

            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "创建时间");
            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区域");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户姓名");

            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

            ExportExcel.createCell(headRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");
            ExportExcel.createCell(headRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if (list != null && list.size() > 0) {
                double totalPraise = 0.0;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuPraiseDetailsEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomerName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getParentBizOrderId());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getWorkCardId());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getStatus().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAreaName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserPhone());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserAddress());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getKeFuName());

                    if(Integer.valueOf(rowData.getStatus().getValue()) >= 40){
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPraiseFee());
                        totalPraise = totalPraise + rowData.getPraiseFee();
                    }else{
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getApplyPraiseFee());
                        totalPraise = totalPraise + rowData.getApplyPraiseFee();
                    }
                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 10));
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPraise);

            }


        } catch (Exception e) {
            log.error("【CustomerPraiseDetailsRptService.customerPraiseDetailsRptExport】客服好评明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
