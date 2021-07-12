package com.kkl.kklplus.provider.rpt.service;


import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

import com.kkl.kklplus.entity.rpt.RPTSpecialChargeAreaEntity;
import com.kkl.kklplus.entity.rpt.common.RPTMiddleTableEnum;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTSpecialChargeSearchCondition;
import com.kkl.kklplus.entity.rpt.web.RPTArea;

import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.SpecialChargeAreaRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.starter.redis.config.RedisGsonService;
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


import javax.annotation.Resource;

import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

import static com.kkl.kklplus.provider.rpt.service.RptBaseService.*;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class SpecialChargeAreaRptService {

    @Resource
    private SpecialChargeAreaRptMapper specialChargeAreaRptMapper;

    @Autowired
    private RedisGsonService redisGsonService;

    @Autowired
    private AreaCacheService areaCacheService;

    /**
     * 获取原表中的数据
     * @param startDate
     * @param endDate
     * @return
     */
    public List<RPTSpecialChargeAreaEntity> getOldSpecialChargeData(Date startDate, Date endDate){
        List<RPTSpecialChargeAreaEntity> oldSpecialDataList = specialChargeAreaRptMapper.getOldSpecialData(startDate, endDate);
        oldSpecialDataList = oldSpecialDataList.stream().filter(i -> !(i.getTravelCharge() == 0 && i.getOtherCharge() == 0)).collect(Collectors.toList());
        int day = Integer.parseInt(DateUtils.getDay(startDate));
        Map<Long, RPTArea> countyMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        for (RPTSpecialChargeAreaEntity entity : oldSpecialDataList) {
                RPTArea rptArea = countyMap.get(entity.getCountyId());
                if (rptArea!= null){
                    RPTArea city = rptArea.getParent();
                    if (city!=null){
                        entity.setCityId(city.getId());
                        RPTArea rptArea1 = cityMap.get(city.getId());
                        RPTArea parent = rptArea1.getParent();
                        if (parent!=null){
                            entity.setProvinceId(parent.getId());
                        }
                    }
                }
                Class countyItemClass = entity.getClass();
                String strSetDMethodName = "setT" + day;
                String strSetAMethodName = "setO" + day;
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Double.class);
                    Method setAMethod = countyItemClass.getMethod(strSetAMethodName, Double.class);
                    setDMethod.invoke(entity, entity.getTravelCharge());
                    setAMethod.invoke(entity, entity.getOtherCharge());
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        return oldSpecialDataList;
    }

    /**
     * 插入数据到中间表
     * @param queryDate
     */
    public void insertSpecialChargeAreaRpt(Date queryDate){
        Date startDate = DateUtils.getDateStart(queryDate);
        Date endDate = DateUtils.getDateEnd(queryDate);
        String yearMonth = DateUtils.getYearMonth(queryDate);
        List<RPTSpecialChargeAreaEntity> oldSpecialChargeData = getOldSpecialChargeData(startDate, endDate);
        int systemId = RptCommonUtils.getSystemId();
        List<Map<String,Long>> countyIds = specialChargeAreaRptMapper.getCountyIds(Integer.parseInt(yearMonth),systemId);
        Map<String,Long> countyMap = new HashMap<>();
        for (Map<String,Long> map:countyIds) {
            countyMap.put(StringUtils.join(map.get("countyId"),":",map.get("productCategoryId")),map.get("id"));
        }
        for (RPTSpecialChargeAreaEntity entity:oldSpecialChargeData) {
            entity.setYearmonth(Integer.valueOf(yearMonth));
            entity.setSystemId(RptCommonUtils.getSystemId());
            String key = StringUtils.join(entity.getCountyId(), ":", entity.getProductCategoryId());
            if (countyMap.get(key)!=null){
                entity.setId(countyMap.get(key));
                specialChargeAreaRptMapper.updateSpecialChargeRpt(entity);
            }else {
                specialChargeAreaRptMapper.insertSpecialChargeRpt(entity);
            }
//
        }
    }


    public List<RPTSpecialChargeAreaEntity> getSpecialChargeNewList(RPTSpecialChargeSearchCondition searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        List<RPTSpecialChargeAreaEntity> specialChargeAreaList = specialChargeAreaRptMapper.getSpecialChargeAreaList(searchCondition.getYearmonth(), searchCondition.getAreaId(),systemId, searchCondition.getProductCategoryIds());
        Map<Long, RPTArea> countyMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        for (RPTSpecialChargeAreaEntity entity:specialChargeAreaList) {
            if (countyMap!=null && countyMap.get(entity.getCountyId())!=null){
                entity.setCountyName(countyMap.get(entity.getCountyId()).getName());
            }
            if (cityMap!=null && cityMap.get(entity.getCityId())!=null){
                entity.setCityName(cityMap.get(entity.getCityId()).getName());
            }
            if (provinceMap!=null && provinceMap.get(entity.getProvinceId())!=null){
                entity.setProvinceName(provinceMap.get(entity.getProvinceId()).getName());
            }
        }
        return specialChargeAreaList;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTSpecialChargeSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSpecialChargeSearchCondition.class);

        if (searchCondition != null && searchCondition.getYearmonth()!=null &&searchCondition.getYearmonth()!=0) {
            Integer rowCount = specialChargeAreaRptMapper.hasReportData(searchCondition.getYearmonth(),
                    searchCondition.getAreaId(),RptCommonUtils.getSystemId(),searchCondition.getProductCategoryIds());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 删除中间表中指定月份的特殊费用数据
     */
    public void deleteSpecialChargeFromRptDB(Integer selectYear,Integer selectMonth) {
        Date date = DateUtils.getDate(selectYear, selectMonth, 1);
        String yearMonth = DateUtils.getYearMonth(date);
        int systemId = RptCommonUtils.getSystemId();
        specialChargeAreaRptMapper.deleteSpecialChargeRpt(systemId, StringUtils.toInteger(yearMonth));

    }

    /**
     *重建中间表
     */
    public boolean rebuildMiddleTableData( RPTRebuildOperationTypeEnum operationType, Integer selectedYear, Integer selectedMonth) {
        boolean result = false;
        if (operationType != null && selectedYear != null && selectedYear > 0
                && selectedMonth != null && selectedMonth > 0) {
            try {
                Date date = DateUtils.getDate(selectedYear, selectedMonth, 1);
                Date beginDate = DateUtils.getStartDayOfMonth(date);
                Date endDate = DateUtils.getLastDayOfMonth(date);
                Date now = new Date();
                if (endDate.getTime() >now.getTime()){
                   endDate =  DateUtils.addDays(now, -1);
                }
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            insertSpecialChargeAreaRpt(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            deleteSpecialChargeFromRptDB(selectedYear,selectedMonth);
                            insertSpecialChargeAreaRpt(beginDate);
                            break;
                        case UPDATE:
                            deleteSpecialChargeFromRptDB(selectedYear,selectedMonth);
                            insertSpecialChargeAreaRpt(beginDate);
                            break;
                        case DELETE:
                            deleteSpecialChargeFromRptDB(selectedYear,selectedMonth);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("SpecialChargeAreaRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 报表导出所需要的数据格式
     * @param searchCondition
     * @param flag
     * @return
     */
    public List<RPTSpecialChargeAreaEntity> getSpecialChargeForExport(RPTSpecialChargeSearchCondition searchCondition,Integer flag){
        List<RPTSpecialChargeAreaEntity> list = getSpecialChargeNewList(searchCondition);

        if (list.size()>0){
            //区县
            Map<Long, List<RPTSpecialChargeAreaEntity>> countyMap = list.stream().collect(Collectors.groupingBy(RPTSpecialChargeAreaEntity::getCountyId));
            List<RPTSpecialChargeAreaEntity> countyList = Lists.newArrayList();
            for (List<RPTSpecialChargeAreaEntity> listItem : countyMap.values()) {
                RPTSpecialChargeAreaEntity countyItem = new RPTSpecialChargeAreaEntity();
                countyItem.setProvinceId(listItem.get(0).getProvinceId());
                countyItem.setProvinceName(listItem.get(0).getProvinceName());
                countyItem.setCityId(listItem.get(0).getCityId());
                countyItem.setCityName(listItem.get(0).getCityName());
                countyItem.setCountyId(listItem.get(0).getCountyId());
                countyItem.setCountyName(listItem.get(0).getCountyName());
                computeSumAndPerForCount(listItem, "O", countyItem);
                computeSumAndPerForCount(listItem, "T", countyItem);
                countyList.add(countyItem);
            }
            //地市
            Map<Long, List<RPTSpecialChargeAreaEntity>> cityMap = Maps.newHashMap();
            for (RPTSpecialChargeAreaEntity item : countyList) {
                List<RPTSpecialChargeAreaEntity> tempCity;
                if (cityMap.containsKey(item.getCityId())) {
                    tempCity = cityMap.get(item.getCityId());
                } else {
                    tempCity = Lists.newArrayList();
                    cityMap.put(item.getCityId(), tempCity);
                }
                tempCity.add(item);
            }
            List<RPTSpecialChargeAreaEntity> cityList = Lists.newArrayList();
            for (List<RPTSpecialChargeAreaEntity> listItem : cityMap.values()) {
                RPTSpecialChargeAreaEntity cityItem = new RPTSpecialChargeAreaEntity();
                cityItem.setProvinceId(listItem.get(0).getProvinceId());
                cityItem.setProvinceName(listItem.get(0).getProvinceName());
                cityItem.setCityId(listItem.get(0).getCityId());
                cityItem.setCityName(listItem.get(0).getCityName());
                cityItem.setCountyId(listItem.get(0).getCountyId());
                cityItem.setCountyName(listItem.get(0).getCountyName());
                cityItem.setItemList(listItem);
                computeSumAndPerForCount(listItem, "O", cityItem);
                computeSumAndPerForCount(listItem, "T", cityItem);
                cityList.add(cityItem);
            }
//            if (flag==2){
//                RPTSpecialChargeAreaEntity sumSCAD = new RPTSpecialChargeAreaEntity();
//                computeSumAndPerForCount(cityList, "O", sumSCAD);
//                computeSumAndPerForCount(cityList, "T", sumSCAD);
//                cityList.add(sumSCAD);
//                return cityList;
//            }
            //省份
            Map<Long, List<RPTSpecialChargeAreaEntity>> provinceMap = Maps.newHashMap();
            for (RPTSpecialChargeAreaEntity item : cityList) {
                List<RPTSpecialChargeAreaEntity> tempProvince;
                if (provinceMap.containsKey(item.getProvinceId())) {
                    tempProvince = provinceMap.get(item.getProvinceId());
                } else {
                    tempProvince = Lists.newArrayList();
                    provinceMap.put(item.getProvinceId(), tempProvince);
                }
                tempProvince.add(item);
            }

            List<RPTSpecialChargeAreaEntity> provinceList = Lists.newArrayList();
            for (List<RPTSpecialChargeAreaEntity> listItem : provinceMap.values()) {
                RPTSpecialChargeAreaEntity provinceItem = new RPTSpecialChargeAreaEntity();

                provinceItem.setProvinceId(listItem.get(0).getProvinceId());
                provinceItem.setProvinceName(listItem.get(0).getProvinceName());
                provinceItem.setCityId(listItem.get(0).getCityId());
                provinceItem.setCityName(listItem.get(0).getCityName());
                provinceItem.setCountyId(listItem.get(0).getCountyId());
                provinceItem.setCountyName(listItem.get(0).getCountyName());

                provinceItem.setItemList(listItem);
                computeSumAndPerForCount(listItem, "O", provinceItem);
                computeSumAndPerForCount(listItem, "T", provinceItem);
                provinceList.add(provinceItem);
            }

            RPTSpecialChargeAreaEntity sumSCAD = new RPTSpecialChargeAreaEntity();
            computeSumAndPerForCount(provinceList, "O", sumSCAD);
            computeSumAndPerForCount(provinceList, "T", sumSCAD);
            provinceList.add(sumSCAD);
            return provinceList;
        }
        return Lists.newArrayList();
    }

    public static void computeSumAndPerForCount(List baseDailyReports,String str,
                                                RPTSpecialChargeAreaEntity sumDailyReport) {

        //计算每日的总单数
        if (sumDailyReport != null) {
            Class sumDailyReportClass = sumDailyReport.getClass();
            for (Object object : baseDailyReports) {
                RPTSpecialChargeAreaEntity item = (RPTSpecialChargeAreaEntity) object;
                Class itemClass = item.getClass();
                for (int i = 1; i < 32; i++) {
                    String strGetMethodName = "get"+str + i;
                    String strSetMethodName = "set"+str + i;
                    try {
                        Method itemGetMethod = itemClass.getMethod(strGetMethodName);
                        Object itemGetD = itemGetMethod.invoke(item);

                        Method sumDailyReportClassGetMethod = sumDailyReportClass.getMethod(strGetMethodName);
                        Object sumDailyReportClassGetD = sumDailyReportClassGetMethod.invoke(sumDailyReport);

                        Double dSum = StringUtils.toDouble(sumDailyReportClassGetD) + StringUtils.toDouble(itemGetD);

                        Method sumDailyReportSetMethod = sumDailyReportClass.getMethod(strSetMethodName, Double.class);
                        sumDailyReportSetMethod.invoke(sumDailyReport, dSum);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }

            }
        }
    }

    /**
     * 省市特殊费用报表导出格式
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportSpecialChargeCityRpt(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTSpecialChargeSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSpecialChargeSearchCondition.class);

            List<RPTSpecialChargeAreaEntity> list = getSpecialChargeForExport(searchCondition,1);

            int days = DateUtils.getDaysOfMonth( DateUtils.parse(String.valueOf(searchCondition.getYearmonth()),"yyyyMM"));

            String xName = reportTitle;

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days * 2 + 3));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 1));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 1 + days * 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 2 + days * 2, 2 + days * 2));
            ExportExcel.createCell(headFirstRow, 2 + days * 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程合计");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 3 + days * 2, 3 + days * 2));
            ExportExcel.createCell(headFirstRow, 3 + days * 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他合计");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                xSheet.addMergedRegion(new CellRangeAddress(2, 2, 1 + (dayIndex * 2 - 1), 1 + (dayIndex * 2)));
                ExportExcel.createCell(headSecondRow, 1 + (dayIndex * 2 - 1), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex);
            }

            Row headThirdRow = xSheet.createRow(rowIndex++);
            headThirdRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headThirdRow, 1 + (dayIndex * 2 - 1), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程");
                ExportExcel.createCell(headThirdRow, 1 + (dayIndex * 2), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //  写入数据
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                int index = 0;
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTSpecialChargeAreaEntity provinceItem = list.get(dataRowIndex);
                    index += 1;
                    if ( index < rowsCount) {
                        //省
                        Row dataRow = xSheet.createRow(rowIndex++);

                        int columnIndex = 0;
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getProvinceName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            String strGetAMethodName = "getO" + dayIndex;

                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Method getAMethod = provinceItemClass.getMethod(strGetAMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);
                            Object objGetA = getAMethod.invoke(provinceItem);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalTravelCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalOtherCharge());


                        //市
                        for (RPTSpecialChargeAreaEntity cityItem : provinceItem.getItemList()) {
                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            columnIndex = 0;

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getProvinceName());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getCityName());

                            Class cityItemClass = cityItem.getClass();
                            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                String strGetDMethodName = "getT" + dayIndex;
                                String strGetAMethodName = "getO" + dayIndex;

                                Method getDMethod = cityItemClass.getMethod(strGetDMethodName);
                                Method getAMethod = cityItemClass.getMethod(strGetAMethodName);
                                Object objGetD = getDMethod.invoke(cityItem);
                                Object objGetA = getAMethod.invoke(cityItem);

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                            }

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getTotalTravelCharge());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getTotalOtherCharge());

                        }

                    } else {
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 2;

                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 1));
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            String strGetAMethodName = "getO" + dayIndex;

                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Method getAMethod = provinceItemClass.getMethod(strGetAMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);
                            Object objGetA = getAMethod.invoke(provinceItem);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalTravelCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalOtherCharge());
                    }

                }

            }
        }catch (Exception e) {
            log.error("省市特殊费用分布报表写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 区、县特殊费用报表导出格式
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportSpecialChargeByCounty(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;

        try {
            RPTSpecialChargeSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSpecialChargeSearchCondition.class);

            List<RPTSpecialChargeAreaEntity> list = getSpecialChargeForExport(searchCondition,1);

            int days = DateUtils.getDaysOfMonth( DateUtils.parse(String.valueOf(searchCondition.getYearmonth()),"yyyyMM"));

            String xName = reportTitle;

            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days * 2 + 4));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 2));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 2 + days * 2));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "费用(元)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 3 + days * 2, 3 + days * 2));
            ExportExcel.createCell(headFirstRow, 3 + days * 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程合计");

            xSheet.addMergedRegion(new CellRangeAddress(1, 3, 4 + days * 2, 4 + days * 2));
            ExportExcel.createCell(headFirstRow, 4 + days * 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他合计");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                xSheet.addMergedRegion(new CellRangeAddress(2, 2, 2 + (dayIndex * 2 - 1), 2 + (dayIndex * 2)));
                ExportExcel.createCell(headSecondRow, 2 + (dayIndex * 2 - 1), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex);
            }

            Row headThirdRow = xSheet.createRow(rowIndex++);
            headThirdRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headThirdRow, 2 + (dayIndex * 2 - 1), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程");
                ExportExcel.createCell(headThirdRow, 2 + (dayIndex * 2), xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //  写入数据
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                int index = 0;
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTSpecialChargeAreaEntity provinceItem = list.get(dataRowIndex);
                    index += 1;
                    if ( index < rowsCount) {
                        //省
                        Row dataRow = xSheet.createRow(rowIndex++);

                        int columnIndex = 0;
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getProvinceName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            String strGetAMethodName = "getO" + dayIndex;

                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Method getAMethod = provinceItemClass.getMethod(strGetAMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);
                            Object objGetA = getAMethod.invoke(provinceItem);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalTravelCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalOtherCharge());


                        //市
                        for (RPTSpecialChargeAreaEntity cityItem : provinceItem.getItemList()) {
                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            columnIndex = 0;

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getProvinceName());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getCityName());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                            Class cityItemClass = cityItem.getClass();
                            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                String strGetDMethodName = "getT" + dayIndex;
                                String strGetAMethodName = "getO" + dayIndex;

                                Method getDMethod = cityItemClass.getMethod(strGetDMethodName);
                                Method getAMethod = cityItemClass.getMethod(strGetAMethodName);
                                Object objGetD = getDMethod.invoke(cityItem);
                                Object objGetA = getAMethod.invoke(cityItem);

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                            }

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getTotalTravelCharge());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getTotalOtherCharge());
                            //区
                            for (RPTSpecialChargeAreaEntity countyItem : cityItem.getItemList()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                columnIndex = 0;

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getProvinceName());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getCityName());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getCountyName());

                                Class countyItemClass = countyItem.getClass();
                                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                    String strGetDMethodName = "getT" + dayIndex;
                                    String strGetAMethodName = "getO" + dayIndex;

                                    Method getDMethod = countyItemClass.getMethod(strGetDMethodName);
                                    Method getAMethod = countyItemClass.getMethod(strGetAMethodName);
                                    Object objGetD = getDMethod.invoke(countyItem);
                                    Object objGetA = getAMethod.invoke(countyItem);

                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                                }

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getTotalTravelCharge());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getTotalOtherCharge());
                            }

                        }

                    } else {
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 3;

                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 2));
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            String strGetAMethodName = "getO" + dayIndex;

                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Method getAMethod = provinceItemClass.getMethod(strGetAMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);
                            Object objGetA = getAMethod.invoke(provinceItem);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetD));
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toDouble(objGetA));
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalTravelCharge());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getTotalOtherCharge());
                    }

                }

            }
        }catch (Exception e) {
            log.error("区县特殊费用分布报表写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
