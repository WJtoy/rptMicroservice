package com.kkl.kklplus.provider.rpt.service;


import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCrushAreaEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTSpecialChargeSearchCondition;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.CrushAreaRptMapper;
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
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static com.kkl.kklplus.provider.rpt.service.RptBaseService.*;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CrushAreaRptService {

    @Resource
    private CrushAreaRptMapper crushAreaRptMapper;

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
    public List<RPTCrushAreaEntity> getCrushAreaData(Date startDate, Date endDate){
        List<RPTCrushAreaEntity> oldSpecialDataList = crushAreaRptMapper.getOldCrushData(startDate, endDate);
        oldSpecialDataList = oldSpecialDataList.stream().filter(i -> !(i.getCrushSum() == 0)).collect(Collectors.toList());
        int day = Integer.parseInt(DateUtils.getDay(startDate));
        for (RPTCrushAreaEntity entity : oldSpecialDataList) {
                Class countyItemClass = entity.getClass();
                String strSetDMethodName = "setT" + day;
                try {
                    Method setDMethod = countyItemClass.getMethod(strSetDMethodName, Integer.class);
                    setDMethod.invoke(entity, entity.getCrushSum());
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
    public void insertCrushAreaRpt(Date queryDate){
        String quarter = QuarterUtils.getSeasonQuarter(queryDate);
        Date startDate = DateUtils.getDateStart(queryDate);
        Date endDate = DateUtils.getDateEnd(queryDate);
        String yearMonth = DateUtils.getYearMonth(queryDate);
        List<RPTCrushAreaEntity> oldSpecialChargeData = getCrushAreaData(startDate, endDate);
        int systemId = RptCommonUtils.getSystemId();
        List<Map<String,Long>> countyIds = crushAreaRptMapper.getCountyIds(Integer.parseInt(yearMonth),systemId);
        Map<String,Long> countyMap = new HashMap<>();
        for (Map<String,Long> map:countyIds) {
            countyMap.put(StringUtils.join(map.get("subAreaId"),":",map.get("productCategoryId")),map.get("id"));
        }
        for (RPTCrushAreaEntity entity:oldSpecialChargeData) {
            entity.setQuarter(quarter);
            entity.setYearmonth(Integer.valueOf(yearMonth));
            entity.setSystemId(RptCommonUtils.getSystemId());
            String key = StringUtils.join(entity.getSubAreaId(), ":", entity.getProductCategoryId());
            if (countyMap.get(key)!=null){
                entity.setId(countyMap.get(key));
                crushAreaRptMapper.updateCrushRpt(entity);
            }else {
                crushAreaRptMapper.insertCrushRpt(entity);
            }

        }
    }


    public List<RPTCrushAreaEntity> getCrushAreaData(RPTSpecialChargeSearchCondition searchCondition){
        int systemId = RptCommonUtils.getSystemId();
        List<RPTCrushAreaEntity> specialChargeAreaList = crushAreaRptMapper.getCrushAreaList(searchCondition.getYearmonth(),systemId, searchCondition.getProductCategoryIds(),searchCondition.getQuarter());
        Map<Long, RPTArea> countyMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long, RPTArea> townMap = areaCacheService.getAllTownMap();
        for (RPTCrushAreaEntity entity:specialChargeAreaList) {
            if (townMap!=null && townMap.get(entity.getSubAreaId())!=null){
                entity.setStreetName(townMap.get(entity.getSubAreaId()).getName());
            }
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
        List<RPTCrushAreaEntity> list = getCrushForExport(specialChargeAreaList);
        return list;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTSpecialChargeSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSpecialChargeSearchCondition.class);

        if (searchCondition != null && searchCondition.getYearmonth()!=null &&searchCondition.getYearmonth()!=0) {
            Integer rowCount = crushAreaRptMapper.hasReportData(searchCondition.getYearmonth(),
                    RptCommonUtils.getSystemId(),searchCondition.getProductCategoryIds(),searchCondition.getQuarter());
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 删除中间表中指定月份的特殊费用数据
     */
    public void deleteCrushFromRptDB(Integer selectYear,Integer selectMonth) {
        Date date = DateUtils.getDate(selectYear, selectMonth, 1);
        String yearMonth = DateUtils.getYearMonth(date);
        int systemId = RptCommonUtils.getSystemId();
        crushAreaRptMapper.deleteCrushRpt(systemId, StringUtils.toInteger(yearMonth));

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
                            insertCrushAreaRpt(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
//                            deleteCrushFromRptDB(selectedYear,selectedMonth);
//                            insertCrushAreaRpt(beginDate);
                            break;
                        case DELETE:
                            deleteCrushFromRptDB(selectedYear,selectedMonth);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("CrushAreaRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 报表导出所需要的数据格式
     * @param
     * @return
     */
    public List<RPTCrushAreaEntity> getCrushForExport(List<RPTCrushAreaEntity> list){

        if (list.size()>0){
            //街道
            Map<Long, List<RPTCrushAreaEntity>> townMap = list.stream().collect(Collectors.groupingBy(RPTCrushAreaEntity::getSubAreaId));
            List<RPTCrushAreaEntity> townList = Lists.newArrayList();
            for (List<RPTCrushAreaEntity> listItem : townMap.values()) {
                RPTCrushAreaEntity townItem = new RPTCrushAreaEntity();
                townItem.setProvinceId(listItem.get(0).getProvinceId());
                townItem.setProvinceName(listItem.get(0).getProvinceName());
                townItem.setCityId(listItem.get(0).getCityId());
                townItem.setCityName(listItem.get(0).getCityName());
                townItem.setCountyId(listItem.get(0).getCountyId());
                townItem.setCountyName(listItem.get(0).getCountyName());
                townItem.setSubAreaId(listItem.get(0).getSubAreaId());
                townItem.setStreetName(listItem.get(0).getStreetName());
                computeSumAndPerForCount(listItem, "T", townItem);
                townList.add(townItem);
            }

            Map<Long, List<RPTCrushAreaEntity>> countyMap = Maps.newHashMap();
            for (RPTCrushAreaEntity item : townList) {
                List<RPTCrushAreaEntity> tempCounty;
                if (countyMap.containsKey(item.getCountyId())) {
                    tempCounty = countyMap.get(item.getCountyId());
                } else {
                    tempCounty = Lists.newArrayList();
                    countyMap.put(item.getCountyId(), tempCounty);
                }
                tempCounty.add(item);
            }

            //区县
            List<RPTCrushAreaEntity> countyList = Lists.newArrayList();
            for (List<RPTCrushAreaEntity> listItem : countyMap.values()) {
                RPTCrushAreaEntity countyItem = new RPTCrushAreaEntity();
                countyItem.setProvinceId(listItem.get(0).getProvinceId());
                countyItem.setProvinceName(listItem.get(0).getProvinceName());
                countyItem.setCityId(listItem.get(0).getCityId());
                countyItem.setCityName(listItem.get(0).getCityName());
                countyItem.setCountyId(listItem.get(0).getCountyId());
                countyItem.setCountyName(listItem.get(0).getCountyName());
                countyItem.setSubAreaId(listItem.get(0).getSubAreaId());
                countyItem.setStreetName(listItem.get(0).getStreetName());
                countyItem.setItemList(listItem);
                computeSumAndPerForCount(listItem, "T", countyItem);
                countyList.add(countyItem);
            }
            //地市
            Map<Long, List<RPTCrushAreaEntity>> cityMap = Maps.newHashMap();
            for (RPTCrushAreaEntity item : countyList) {
                List<RPTCrushAreaEntity> tempCity;
                if (cityMap.containsKey(item.getCityId())) {
                    tempCity = cityMap.get(item.getCityId());
                } else {
                    tempCity = Lists.newArrayList();
                    cityMap.put(item.getCityId(), tempCity);
                }
                tempCity.add(item);
            }
            List<RPTCrushAreaEntity> cityList = Lists.newArrayList();
            for (List<RPTCrushAreaEntity> listItem : cityMap.values()) {
                RPTCrushAreaEntity cityItem = new RPTCrushAreaEntity();
                cityItem.setProvinceId(listItem.get(0).getProvinceId());
                cityItem.setProvinceName(listItem.get(0).getProvinceName());
                cityItem.setCityId(listItem.get(0).getCityId());
                cityItem.setCityName(listItem.get(0).getCityName());
                cityItem.setCountyId(listItem.get(0).getCountyId());
                cityItem.setCountyName(listItem.get(0).getCountyName());
                cityItem.setItemList(listItem);
                computeSumAndPerForCount(listItem, "T", cityItem);
                cityList.add(cityItem);
            }
            //省份
            Map<Long, List<RPTCrushAreaEntity>> provinceMap = Maps.newHashMap();
            for (RPTCrushAreaEntity item : cityList) {
                List<RPTCrushAreaEntity> tempProvince;
                if (provinceMap.containsKey(item.getProvinceId())) {
                    tempProvince = provinceMap.get(item.getProvinceId());
                } else {
                    tempProvince = Lists.newArrayList();
                    provinceMap.put(item.getProvinceId(), tempProvince);
                }
                tempProvince.add(item);
            }

            List<RPTCrushAreaEntity> provinceList = Lists.newArrayList();
            for (List<RPTCrushAreaEntity> listItem : provinceMap.values()) {
                RPTCrushAreaEntity provinceItem = new RPTCrushAreaEntity();

                provinceItem.setProvinceId(listItem.get(0).getProvinceId());
                provinceItem.setProvinceName(listItem.get(0).getProvinceName());
                provinceItem.setCityId(listItem.get(0).getCityId());
                provinceItem.setCityName(listItem.get(0).getCityName());
                provinceItem.setCountyId(listItem.get(0).getCountyId());
                provinceItem.setCountyName(listItem.get(0).getCountyName());

                provinceItem.setItemList(listItem);
                computeSumAndPerForCount(listItem, "T", provinceItem);
                provinceList.add(provinceItem);
            }

            RPTCrushAreaEntity sumSCAD = new RPTCrushAreaEntity();
            computeSumAndPerForCount(provinceList, "T", sumSCAD);
            provinceList.add(sumSCAD);
            return provinceList;
        }
        return Lists.newArrayList();
    }

    public static void computeSumAndPerForCount(List baseDailyReports,String str,
                                                RPTCrushAreaEntity sumDailyReport) {

        //计算每日的总单数
        if (sumDailyReport != null) {
            Class sumDailyReportClass = sumDailyReport.getClass();
            for (Object object : baseDailyReports) {
                RPTCrushAreaEntity item = (RPTCrushAreaEntity) object;
                Class itemClass = item.getClass();
                for (int i = 1; i < 32; i++) {
                    String strGetMethodName = "get"+str + i;
                    String strSetMethodName = "set"+str + i;
                    try {
                        Method itemGetMethod = itemClass.getMethod(strGetMethodName);
                        Object itemGetD = itemGetMethod.invoke(item);

                        Method sumDailyReportClassGetMethod = sumDailyReportClass.getMethod(strGetMethodName);
                        Object sumDailyReportClassGetD = sumDailyReportClassGetMethod.invoke(sumDailyReport);

                        Integer dSum = StringUtils.toInteger(sumDailyReportClassGetD) + StringUtils.toInteger(itemGetD);

                        Method sumDailyReportSetMethod = sumDailyReportClass.getMethod(strSetMethodName, Integer.class);
                        sumDailyReportSetMethod.invoke(sumDailyReport, dSum);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }

            }
        }
    }


    /**
     * 突击单量报表导出格式
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportCrushArea(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;

        try {
            RPTSpecialChargeSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSpecialChargeSearchCondition.class);

            List<RPTCrushAreaEntity> list = getCrushAreaData(searchCondition);

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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 4));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 3));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区/街道");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 4, 3 + days));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "突击单量(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4 + days, 4 + days));
            ExportExcel.createCell(headFirstRow, 4 + days, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, 3 + dayIndex, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER,dayIndex );
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //  写入数据
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                int index = 0;
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTCrushAreaEntity provinceItem = list.get(dataRowIndex);
                    index += 1;
                    if ( index < rowsCount) {
                        //省
                        Row dataRow = xSheet.createRow(rowIndex++);

                        int columnIndex = 0;
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getProvinceName());
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toInteger(objGetD));

                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getCrushTotal());


                        //市
                        for (RPTCrushAreaEntity cityItem : provinceItem.getItemList()) {
                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            columnIndex = 0;

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getProvinceName());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getCityName());
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                            Class cityItemClass = cityItem.getClass();
                            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                String strGetDMethodName = "getT" + dayIndex;

                                Method getDMethod = cityItemClass.getMethod(strGetDMethodName);
                                Object objGetD = getDMethod.invoke(cityItem);

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toInteger(objGetD));
                            }

                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityItem.getCrushTotal());
                            //区
                            for (RPTCrushAreaEntity countyItem : cityItem.getItemList()) {
                                dataRow = xSheet.createRow(rowIndex++);
                                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                columnIndex = 0;

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getProvinceName());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getCityName());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getCountyName());
                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);

                                Class countyItemClass = countyItem.getClass();
                                for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                    String strGetDMethodName = "getT" + dayIndex;
                                    Method getDMethod = countyItemClass.getMethod(strGetDMethodName);
                                    Object objGetD = getDMethod.invoke(countyItem);


                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toInteger(objGetD));

                                }

                                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, countyItem.getCrushTotal());
                                //街道
                                for(RPTCrushAreaEntity townItem : countyItem.getItemList()){
                                    dataRow = xSheet.createRow(rowIndex++);
                                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                    columnIndex = 0;

                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townItem.getProvinceName());
                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townItem.getCityName());
                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townItem.getCountyName());
                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townItem.getStreetName());

                                    Class townItemClass = townItem.getClass();
                                    for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                                        String strGetDMethodName = "getT" + dayIndex;
                                        Method getDMethod = townItemClass.getMethod(strGetDMethodName);
                                        Object objGetD = getDMethod.invoke(townItem);


                                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toInteger(objGetD));

                                    }

                                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townItem.getCrushTotal());
                                }
                            }

                        }

                    } else {
                        Row dataRow = xSheet.createRow(rowIndex++);
                        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                        int columnIndex = 4;

                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 3));
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");

                        Class provinceItemClass = provinceItem.getClass();
                        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                            String strGetDMethodName = "getT" + dayIndex;
                            Method getDMethod = provinceItemClass.getMethod(strGetDMethodName);
                            Object objGetD = getDMethod.invoke(provinceItem);


                            ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, StringUtils.toInteger(objGetD));
                        }

                        ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, provinceItem.getCrushTotal());

                    }

                }

            }
        }catch (Exception e) {
            log.error("突击单量分布报表写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
