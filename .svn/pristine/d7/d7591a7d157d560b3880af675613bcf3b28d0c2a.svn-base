package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTKeFuAverageOrderFeeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.KeFuAverageOrderFeeRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.utils.StringUtils;
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

import java.text.NumberFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuAverageOrderFeeRptService extends RptBaseService {

    @Autowired
    private KeFuAverageOrderFeeRptMapper keFuAverageOrderFeeRptMapper;


    @Autowired
    private AreaCacheService areaCacheService;

    @Autowired
    private MSCustomerService customerService;

    public List<RPTKeFuAverageOrderFeeEntity> getKeFuAverageOrderFee(RPTComplainStatisticsDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int systemId = RptCommonUtils.getSystemId();
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        List<RPTKeFuAverageOrderFeeEntity> list  = new ArrayList<>();
        List<RPTKeFuAverageOrderFeeEntity> specialList = keFuAverageOrderFeeRptMapper.getSpecialOrderFeeList(systemId,yearMonth,search.getQuarter(),search.getProductCategoryIds(),search.getCustomerId(),search.getAreaType(),search.getAreaId());
        List<RPTKeFuAverageOrderFeeEntity> gradeOrderList = keFuAverageOrderFeeRptMapper.getGradedOrderData(systemId,yearMonth,search.getQuarter(),search.getProductCategoryIds(),search.getCustomerId(),search.getAreaType(),search.getAreaId());
        List<Long> vipCustomerList = keFuAverageOrderFeeRptMapper.getVipCustomer(systemId);

        specialList = specialList.stream().filter(x -> !vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());

        gradeOrderList = gradeOrderList.stream().filter(x -> !vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());


        Map<String, List<RPTKeFuAverageOrderFeeEntity>> specialMap = specialList.stream().collect(Collectors.groupingBy(e -> fetchGroupKey(e)));
        Map<String, List<RPTKeFuAverageOrderFeeEntity>> gradeOrderMap = gradeOrderList.stream().collect(Collectors.groupingBy(e -> fetchGroupKey(e)));

        Map<Long,RPTArea>  cityMap = areaCacheService.getAllCityMap();
        Map<Long,RPTArea>  provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long,RPTArea>  areaMap = areaCacheService.getAllCountyMap();

        RPTKeFuAverageOrderFeeEntity entity;
        Long cityId = 0l;
        Long provinceId = 0l;
        Long areaId;
        for( RPTArea rptArea : areaMap.values()){
            Double orderFee = 0.0;
            Double travelFee = 0.0;
            Integer orderSum = 0;
            String cityName = "";
            String provinceName = "";
            String countyName = "";
            entity = new RPTKeFuAverageOrderFeeEntity();
            countyName = rptArea.getName();
            areaId = rptArea.getId();
            entity.setAreaId(areaId);
            RPTArea city = rptArea.getParent();
            if (city!=null){
                cityId = city.getId();
                entity.setCityId(cityId);
                RPTArea cityArea = cityMap.get(cityId);
                if (cityArea!=null){
                    cityName = cityArea.getName();
                    RPTArea province =  cityArea.getParent();
                    if(province != null){
                        provinceId =  province.getId();
                        RPTArea provinceArea = provinceMap.get(provinceId);
                        entity.setProvinceId(provinceId);
                        if (provinceArea!=null){
                            provinceName = provinceArea.getName();
                        }
                    }
                }
            }
            entity.setAreaName(provinceName+cityName+countyName);
            String key = StringUtils.join(provinceId,"%",cityId,"%",areaId);
            List<RPTKeFuAverageOrderFeeEntity>  gradeList = gradeOrderMap.get(key);

            List<RPTKeFuAverageOrderFeeEntity>  orderFeeList = specialMap.get(key);


            //远程和其他费用
            if(orderFeeList !=null && !orderFeeList.isEmpty() ) {
                for (RPTKeFuAverageOrderFeeEntity item : orderFeeList) {
                    travelFee += item.getTravelFee();
                    orderFee += item.getOrderFee();
                }
            }

            //完工单数量
            if (gradeList != null && !gradeList.isEmpty()) {
                    for (RPTKeFuAverageOrderFeeEntity keFuAverageOrderFeeEntity : gradeList) {
                        orderSum += keFuAverageOrderFeeEntity.getOrderSum();
                    }
                }


            entity.setTravelFee(travelFee);
            entity.setOrderFee(orderFee);
            entity.setOrderSum(orderSum);
            entity.setAverageOrderFee(orderSum > 0 ? (travelFee+orderFee)/orderSum : 0);
            if(entity.getProvinceId() != null){
                list.add(entity);
            }

        }

        if(null != search.getAreaId()  && search.getAreaId() != 0){
            if(search.getAreaType() == 4){
                list = list.stream().filter(x -> x.getAreaId().equals(search.getAreaId())).collect(Collectors.toList());
            }else if(search.getAreaType() == 3){
                list = list.stream().filter(x -> x.getCityId().equals(search.getAreaId())).collect(Collectors.toList());
            }else if(search.getAreaType() == 2){
                list = list.stream().filter(x -> x.getProvinceId().equals(search.getAreaId())).collect(Collectors.toList());
            }
        }

        if(null != search.getKeFuId() && search.getKeFuId() != 0){
            Integer row = keFuAverageOrderFeeRptMapper.getKeFuAreaC(systemId,search.getKeFuId());
            if(row == 0){
                List<RPTKeFuAverageOrderFeeEntity> keFuA =  keFuAverageOrderFeeRptMapper.getKeFuAreaA(systemId,search.getKeFuId());

                List<RPTKeFuAverageOrderFeeEntity> keFuB = keFuAverageOrderFeeRptMapper.getKeFuAreaB(systemId,search.getKeFuId());

                List<RPTKeFuAverageOrderFeeEntity> keFuD = keFuAverageOrderFeeRptMapper.getKeFuAreaD(systemId,search.getKeFuId());

                List<Long> keFuListAreaId = keFuD.stream().map(RPTKeFuAverageOrderFeeEntity::getAreaId).distinct().collect(Collectors.toList());

                List<Long> keFuListCityId = keFuA.stream().map(RPTKeFuAverageOrderFeeEntity::getCityId).distinct().collect(Collectors.toList());

                List<Long> keFuListProvinceId = keFuB.stream().map(RPTKeFuAverageOrderFeeEntity::getProvinceId).distinct().collect(Collectors.toList());

                list = list.stream().filter(x -> keFuListCityId.contains(x.getCityId()) ||keFuListProvinceId.contains(x.getProvinceId()) ||keFuListAreaId.contains(x.getAreaId())).collect(Collectors.toList());
            }

        }
        list = list.stream().sorted(Comparator.comparing(RPTKeFuAverageOrderFeeEntity::getProvinceId)).collect(Collectors.toList());
        return list;
    }



    public List<RPTKeFuAverageOrderFeeEntity> getKAKeFuAverageOrderFee(RPTComplainStatisticsDailySearch search) {
        Integer selectedYear = DateUtils.getYear(new Date(search.getStartDate()));
        Integer selectedMonth = DateUtils.getMonth(new Date(search.getStartDate()));
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int systemId = RptCommonUtils.getSystemId();
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        List<RPTKeFuAverageOrderFeeEntity> list  = new ArrayList<>();
        List<RPTKeFuAverageOrderFeeEntity> specialList = keFuAverageOrderFeeRptMapper.getSpecialOrderFeeList(systemId,yearMonth,search.getQuarter(),search.getProductCategoryIds(),search.getCustomerId(),search.getAreaType(),search.getAreaId());
        List<RPTKeFuAverageOrderFeeEntity> gradeOrderList = keFuAverageOrderFeeRptMapper.getGradedOrderData(systemId,yearMonth,search.getQuarter(),search.getProductCategoryIds(),search.getCustomerId(),search.getAreaType(),search.getAreaId());
        List<Long> vipCustomerList = keFuAverageOrderFeeRptMapper.getVipCustomer(systemId);

        if(search.getKeFuId() != null && search.getKeFuId() != 0){
            List<Long> keFuCustomerList = keFuAverageOrderFeeRptMapper.getKeFuCustomer(search.getKeFuId(),systemId);
            specialList = specialList.stream().filter(x -> keFuCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());
            gradeOrderList = gradeOrderList.stream().filter(x -> keFuCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());

        }

        specialList = specialList.stream().filter(x -> vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());


        gradeOrderList = gradeOrderList.stream().filter(x -> vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());

        Map<String, List<RPTKeFuAverageOrderFeeEntity>> specialMap = specialList.stream().collect(Collectors.groupingBy(e -> fetchGroupKey(e)));
        Map<String, List<RPTKeFuAverageOrderFeeEntity>> gradeOrderMap = gradeOrderList.stream().collect(Collectors.groupingBy(e -> fetchGroupKey(e)));

        Map<Long,RPTArea>  cityMap = areaCacheService.getAllCityMap();
        Map<Long,RPTArea>  provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long,RPTArea>  areaMap = areaCacheService.getAllCountyMap();

        RPTKeFuAverageOrderFeeEntity entity;
        Long areaId;
        Long cityId = 0l;
        Long provinceId = 0l;
        for( RPTArea rptArea : areaMap.values()){
            Double orderFee = 0.0;
            Double travelFee = 0.0;
            Integer orderSum = 0;
            String cityName = "";
            String provinceName = "";
            String countyName = "";
            entity = new RPTKeFuAverageOrderFeeEntity();
            countyName = rptArea.getName();
            areaId = rptArea.getId();
            entity.setAreaId(areaId);
            RPTArea city = rptArea.getParent();
            if (city!=null){
                cityId = city.getId();
                entity.setCityId(cityId);
                RPTArea cityArea = cityMap.get(cityId);
                if (cityArea!=null){
                    cityName = cityArea.getName();
                    RPTArea province =  cityArea.getParent();
                    if(province != null){
                        provinceId =  province.getId();
                        RPTArea provinceArea = provinceMap.get(provinceId);
                        entity.setProvinceId(provinceId);
                        if (provinceArea!=null){
                            provinceName = provinceArea.getName();
                        }
                    }
                }
            }
            entity.setAreaName(provinceName+cityName+countyName);
            String key = StringUtils.join(provinceId,"%",cityId,"%",areaId);
            List<RPTKeFuAverageOrderFeeEntity>  gradeList = gradeOrderMap.get(key);

            List<RPTKeFuAverageOrderFeeEntity>  orderFeeList = specialMap.get(key);


            //远程和其他费用
            if(orderFeeList !=null && !orderFeeList.isEmpty() ) {
                for (RPTKeFuAverageOrderFeeEntity item : orderFeeList) {
                    travelFee += item.getTravelFee();
                    orderFee += item.getOrderFee();
                }
            }

            //完工单数量
            if(gradeList != null && !gradeList.isEmpty()){
                for(RPTKeFuAverageOrderFeeEntity keFuAverageOrderFeeEntity : gradeList){
                    orderSum += keFuAverageOrderFeeEntity.getOrderSum();
                }
            }

            entity.setTravelFee(travelFee);
            entity.setOrderFee(orderFee);
            entity.setOrderSum(orderSum);
            entity.setAverageOrderFee(orderSum > 0 ? (travelFee+orderFee)/orderSum : 0);
            if(entity.getProvinceId() != null){
                list.add(entity);
            }
        }

        if(null != search.getAreaId()  && search.getAreaId() != 0){
            if(search.getAreaType() == 4){
                list = list.stream().filter(x -> x.getAreaId().equals(search.getAreaId())).collect(Collectors.toList());
            }else if(search.getAreaType() == 3){
                list = list.stream().filter(x -> x.getCityId().equals(search.getAreaId())).collect(Collectors.toList());
            }else if(search.getAreaType() == 2){
                list = list.stream().filter(x -> x.getProvinceId().equals(search.getAreaId())).collect(Collectors.toList());
            }
        }

        if(null != search.getKeFuId()  && search.getKeFuId() != 0){
            Integer row = keFuAverageOrderFeeRptMapper.getKeFuAreaC(systemId,search.getKeFuId());
            if(row == 0){
                List<RPTKeFuAverageOrderFeeEntity> keFuA =  keFuAverageOrderFeeRptMapper.getKeFuAreaA(systemId,search.getKeFuId());

                List<RPTKeFuAverageOrderFeeEntity> keFuB = keFuAverageOrderFeeRptMapper.getKeFuAreaB(systemId,search.getKeFuId());

                List<RPTKeFuAverageOrderFeeEntity> keFuD = keFuAverageOrderFeeRptMapper.getKeFuAreaD(systemId,search.getKeFuId());

                List<Long> keFuListAreaId = keFuD.stream().map(RPTKeFuAverageOrderFeeEntity::getAreaId).distinct().collect(Collectors.toList());

                List<Long> keFuListCityId = keFuA.stream().map(RPTKeFuAverageOrderFeeEntity::getCityId).distinct().collect(Collectors.toList());

                List<Long> keFuListProvinceId = keFuB.stream().map(RPTKeFuAverageOrderFeeEntity::getProvinceId).distinct().collect(Collectors.toList());

                list = list.stream().filter(x -> keFuListCityId.contains(x.getCityId()) ||keFuListProvinceId.contains(x.getProvinceId()) ||keFuListAreaId.contains(x.getAreaId())).collect(Collectors.toList());
            }

        }
        list = list.stream().sorted(Comparator.comparing(RPTKeFuAverageOrderFeeEntity::getProvinceId)).collect(Collectors.toList());
        return list;
    }

    private static String fetchGroupKey(RPTKeFuAverageOrderFeeEntity entity){
        return StringUtils.join(entity.getProvinceId() ,"%" ,entity.getCityId(),"%",entity.getAreaId());
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
        Date beginDate = new Date(searchCondition.getStartDate());
        if (new Date().getTime() < beginDate.getTime()) {
            return false;
        }
        Integer selectedYear = DateUtils.getYear(beginDate);
        Integer selectedMonth = DateUtils.getMonth(beginDate);
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int systemId = RptCommonUtils.getSystemId();
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        List<RPTKeFuAverageOrderFeeEntity> rptKeFuAverageOrderFeeEntities = keFuAverageOrderFeeRptMapper.getOrderSum(systemId,yearMonth,searchCondition.getQuarter(),searchCondition.getProductCategoryIds(),searchCondition.getCustomerId(),searchCondition.getAreaType(),searchCondition.getAreaId());
        List<Long> vipCustomerList = keFuAverageOrderFeeRptMapper.getVipCustomer(systemId);
        rptKeFuAverageOrderFeeEntities = rptKeFuAverageOrderFeeEntities.stream().filter(x -> !vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());

        result = rptKeFuAverageOrderFeeEntities.size() > 0;
        return result;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasVipReportData(String searchConditionJson) {
        boolean result = true;
        RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
        Date beginDate = new Date(searchCondition.getStartDate());
        if (new Date().getTime() < beginDate.getTime()) {
            return false;
        }
        Integer selectedYear = DateUtils.getYear(beginDate);
        Integer selectedMonth = DateUtils.getMonth(beginDate);
        Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
        int systemId = RptCommonUtils.getSystemId();
        Integer yearMonth = StringUtils.toInteger(DateUtils.formatDate(queryDate, "yyyyMM"));
        List<RPTKeFuAverageOrderFeeEntity> rptKeFuAverageOrderFeeEntities = keFuAverageOrderFeeRptMapper.getVipOrderSum(systemId,yearMonth,searchCondition.getQuarter(),searchCondition.getProductCategoryIds(),searchCondition.getCustomerId(),searchCondition.getAreaType(),searchCondition.getAreaId(),searchCondition.getKeFuId());
        List<Long> vipCustomerList = keFuAverageOrderFeeRptMapper.getVipCustomer(systemId);
        rptKeFuAverageOrderFeeEntities = rptKeFuAverageOrderFeeEntities.stream().filter(x -> vipCustomerList.contains(x.getCustomerId())).collect(Collectors.toList());
        Integer rowCount = rptKeFuAverageOrderFeeEntities.size();
        result = rowCount > 0;
        return result;
    }





    public SXSSFWorkbook keFuAverageOrderFeeRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
            List<RPTKeFuAverageOrderFeeEntity> list = getKeFuAverageOrderFee(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成单数");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费用");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "均单费用(元)");

            rowIndex ++;
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if(list != null && list.size() > 0){
                Integer totalOrderSum = 0;
                double totalOrderFee = 0.0;
                double totalTravelFee = 0.0;
                double totalFee = 0.0;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuAverageOrderFeeEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAreaName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderSum());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTravelFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTravelFee()+rowData.getOrderFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f",rowData.getAverageOrderFee()));
                    totalOrderSum = totalOrderSum + rowData.getOrderSum();
                    totalTravelFee = totalTravelFee + rowData.getTravelFee();
                    totalOrderFee = totalOrderFee + rowData.getOrderFee();
                    totalFee = totalFee + rowData.getOrderFee()+rowData.getTravelFee();


                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,totalOrderSum);
                ExportExcel.createCell(sumRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTravelFee);
                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOrderFee);
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,totalFee);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f",totalOrderSum == 0 ?totalFee : totalFee/totalOrderSum));

            }

        } catch (Exception e) {
            log.error("【KeFuAverageOrderFeeRptService.KeFuAverageOrderFeeRptExport】客服均单费用写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }



    public SXSSFWorkbook vipKeFuAverageOrderFeeRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
            List<RPTKeFuAverageOrderFeeEntity> list = getKAKeFuAverageOrderFee(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成单数");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费用");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "均单费用(元)");


            rowIndex ++;
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if(list != null && list.size() > 0){
                Integer totalOrderSum = 0;
                double totalOrderFee = 0.0;
                double totalTravelFee = 0.0;
                double totalFee = 0.0;

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTKeFuAverageOrderFeeEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAreaName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderSum());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTravelFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTravelFee()+rowData.getOrderFee());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f",rowData.getAverageOrderFee()));
                    totalOrderSum = totalOrderSum + rowData.getOrderSum();
                    totalOrderFee = totalOrderFee + rowData.getOrderFee();
                    totalTravelFee = totalTravelFee + rowData.getTravelFee();
                    totalFee = totalFee + rowData.getOrderFee()+rowData.getTravelFee();

                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOrderSum);
                ExportExcel.createCell(sumRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTravelFee);
                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOrderFee);
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalFee);
                ExportExcel.createCell(sumRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, String.format("%.2f",totalOrderSum == 0 ?totalFee : totalFee/totalOrderSum));

            }

        } catch (Exception e) {
            log.error("【KeFuAverageOrderFeeRptService.vipKeFuAverageOrderFeeRptExport】KA均单费用写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


}
