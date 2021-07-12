package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTAreaOrderPlanDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTBaseDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTAreaOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTSearchBase;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.AreaOrderPlanDailyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
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
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class AreaOrderPlanDailyRptService extends RptBaseService {
    @Resource
    private AreaOrderPlanDailyRptMapper areaOrderPlanDailyRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;

    /**
     * 根据条件查询省市区每日下单单量
     */
    public Map<String, List<RPTAreaOrderPlanDailyEntity>> getAreaOrderPlanDailyRptData(RPTAreaOrderPlanDailySearch search) {

        int systemId = RptCommonUtils.getSystemId();
        search.setSystemId(systemId);
        List<RPTAreaOrderPlanDailyEntity> provinceList = new ArrayList<>();

        List<RPTAreaOrderPlanDailyEntity> cityList = new ArrayList<>();

        List<RPTAreaOrderPlanDailyEntity> areaList = new ArrayList<>();


        List<RPTAreaOrderPlanDailyEntity> provinceOrderPlanDay = areaOrderPlanDailyRptMapper.getProvinceOrderPlanDay(search);


        List<RPTAreaOrderPlanDailyEntity> cityOrderPlanDay = areaOrderPlanDailyRptMapper.getCityOrderPlanDay(search);


        List<RPTAreaOrderPlanDailyEntity> areaOrderPlanDay = areaOrderPlanDailyRptMapper.getAreaOrderPlanDay(search);


        Map<Long, List<RPTAreaOrderPlanDailyEntity>> provincePlanMap = provinceOrderPlanDay.stream().collect(Collectors.groupingBy(RPTAreaOrderPlanDailyEntity::getProvinceId));

        Map<Long, List<RPTAreaOrderPlanDailyEntity>> cityPlanMap = cityOrderPlanDay.stream().collect(Collectors.groupingBy(RPTAreaOrderPlanDailyEntity::getCityId));

        Map<Long, List<RPTAreaOrderPlanDailyEntity>> areaPlanMap = areaOrderPlanDay.stream().collect(Collectors.groupingBy(RPTAreaOrderPlanDailyEntity::getAreaId));

        double total;
        RPTArea province;
        RPTArea city;
        RPTArea area;
        String provinceName;
        String cityName;
        String areaName;
        Long Id;
        Long provinceId;
        Long cityId;
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        RPTAreaOrderPlanDailyEntity rptEntity;
        try {
            for (List<RPTAreaOrderPlanDailyEntity> entity : provincePlanMap.values()) {
                rptEntity = new RPTAreaOrderPlanDailyEntity();
                total = 0;
                provinceId = entity.get(0).getProvinceId();
                province = provinceMap.get(provinceId);
                if (province != null) {
                    provinceName = province.getName();
                    rptEntity.setProvinceId(provinceId);
                    rptEntity.setProvinceName(provinceName);
                }
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);
                rptEntity.setTotal(total);
                provinceList.add(rptEntity);
            }

            for (List<RPTAreaOrderPlanDailyEntity> entity : cityPlanMap.values()) {
                rptEntity = new RPTAreaOrderPlanDailyEntity();
                total = 0;
                provinceId = entity.get(0).getProvinceId();
                cityId = entity.get(0).getCityId();
                city = cityMap.get(cityId);
                if (city != null) {
                    cityName = city.getName();
                    rptEntity.setProvinceId(provinceId);
                    rptEntity.setCityId(cityId);
                    rptEntity.setCityName(cityName);
                }
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);
                rptEntity.setTotal(total);
                cityList.add(rptEntity);
            }

            for (List<RPTAreaOrderPlanDailyEntity> entity : areaPlanMap.values()) {
                rptEntity = new RPTAreaOrderPlanDailyEntity();
                total = 0;
                provinceId = entity.get(0).getProvinceId();
                cityId = entity.get(0).getCityId();
                Id = entity.get(0).getAreaId();
                province = provinceMap.get(provinceId);
                city = cityMap.get(cityId);
                area = areaMap.get(Id);
                if (province != null) {
                    provinceName = province.getName();
                    rptEntity.setProvinceId(provinceId);
                    rptEntity.setProvinceName(provinceName);
                }
                if (city != null) {
                    rptEntity.setCityId(cityId);
                    cityName = city.getName();
                    rptEntity.setCityName(cityName);
                }
                if (area != null) {
                    rptEntity.setAreaId(Id);
                    areaName = area.getName();
                    rptEntity.setAreaName(areaName);
                }
                Class rptEntityClass = rptEntity.getClass();
                total = writeDailyOrders(rptEntity, total, rptEntityClass, entity);
                rptEntity.setTotal(total);
                areaList.add(rptEntity);
            }

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        RPTAreaOrderPlanDailyEntity sumUp = new RPTAreaOrderPlanDailyEntity();
        sumUp.setProvinceId(-1L);
        sumUp.setProvinceName("总计(单)");
        RPTBaseDailyEntity.computeSumAndPerForCount(areaList, 0, 0, sumUp, null);

        areaList = areaList.stream().sorted(Comparator.comparing(RPTAreaOrderPlanDailyEntity::getProvinceId)
                                                    .thenComparing(RPTAreaOrderPlanDailyEntity :: getCityId)).collect(Collectors.toList());

        Map<String, List<RPTAreaOrderPlanDailyEntity>> map = Maps.newHashMap();
        List<RPTAreaOrderPlanDailyEntity> sumUpList = new ArrayList<>();
        sumUpList.add(sumUp);
        map.put(RPTAreaOrderPlanDailyEntity.MAP_KEY_PROVINCELIST, provinceList);
        map.put(RPTAreaOrderPlanDailyEntity.MAP_KEY_CITYLIST, cityList);
        map.put(RPTAreaOrderPlanDailyEntity.MAP_KEY_AREALIST, areaList);
        map.put(RPTAreaOrderPlanDailyEntity.MAP_KEY_SUMUP, sumUpList);
        return map;
    }


    /**
     * 写入每日的订单
     *
     * @param rptEntity
     * @param total
     * @param rptEntityClass
     * @param entityList
     * @return
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private double writeDailyOrders(RPTAreaOrderPlanDailyEntity rptEntity, double total, Class rptEntityClass, List<RPTAreaOrderPlanDailyEntity> entityList) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        int dayIndex;
        String dateStr;
        int day;
        String strSetMethodName;
        Method sumDailyReportSetMethod;
        double daySum;
        for (RPTAreaOrderPlanDailyEntity entity : entityList) {
            dayIndex = entity.getDayIndex();
            daySum = entity.getOrderSum();
            if (dayIndex != 0) {
                dateStr = String.valueOf(dayIndex);
                day = StringUtils.toInteger(dateStr);
                strSetMethodName = "setD" + day;
                sumDailyReportSetMethod = rptEntityClass.getMethod(strSetMethodName, Double.class);
                total += daySum;
                sumDailyReportSetMethod.invoke(rptEntity, daySum);
            }

        }
        return total;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTAreaOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTAreaOrderPlanDailySearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getStartDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = areaOrderPlanDailyRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 省市区每日下单列表导出
     *
     * @return
     */
    public SXSSFWorkbook areaOrderPlanDayRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTAreaOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTAreaOrderPlanDailySearch.class);
            Map<String, List<RPTAreaOrderPlanDailyEntity>> entityMap = getAreaOrderPlanDailyRptData(searchCondition);
            int days = DateUtils.getDaysOfMonth(new Date(searchCondition.getStartDate()));
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, days + 3));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, days + 2));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "每日下单(单)");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, days + 3, days + 3));
            ExportExcel.createCell(headFirstRow, days + 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计(单)");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
                ExportExcel.createCell(headSecondRow, dayIndex + 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, dayIndex + "");
            }

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            List<RPTAreaOrderPlanDailyEntity> pList = entityMap.get(RPTAreaOrderPlanDailyEntity.MAP_KEY_PROVINCELIST);
            List<RPTAreaOrderPlanDailyEntity> cList = entityMap.get(RPTAreaOrderPlanDailyEntity.MAP_KEY_CITYLIST);
            List<RPTAreaOrderPlanDailyEntity> aList = entityMap.get(RPTAreaOrderPlanDailyEntity.MAP_KEY_AREALIST);
            RPTAreaOrderPlanDailyEntity sumAOPD = entityMap.get(RPTAreaOrderPlanDailyEntity.MAP_KEY_SUMUP).get(0);

            // 写入数据
            Row dataRow = null;
            Cell dataCell = null;
            if (pList != null && pList.size() > 0) {
                int pCount = pList.size();
                // 循环读取所有的省
                for (int i = 0; i < pCount; i++) {
                    RPTAreaOrderPlanDailyEntity province = pList.get(i);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int pColumnIndex = 0;

                    dataCell = ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(null == province.getProvinceName() ? "" : province.getProvinceName());

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                    Class provinceClass = province.getClass();
                    pColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, province, pColumnIndex, provinceClass);

                    Double pTotalCount = StringUtils.toDouble(province.getTotal());
                    String strPTotalCount = String.format("%.0f", pTotalCount);

                    ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strPTotalCount);

                    //循环读取省下的市
                    for (RPTAreaOrderPlanDailyEntity city : cList) {
                        if (city.getProvinceId().equals(province.getProvinceId())) {

                            dataRow = xSheet.createRow(rowIndex++);
                            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                            int cColumnIndex = 0;

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                            dataCell = ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                            dataCell.setCellValue(null == city.getCityName() ? "" : city.getCityName());

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                            Class cityClass = city.getClass();

                            cColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, city, cColumnIndex, cityClass);

                            Double cTotalCount = StringUtils.toDouble(city.getTotal());
                            String strCTotalCount = String.format("%.0f", cTotalCount);

                            ExportExcel.createCell(dataRow, cColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strCTotalCount);

                            //循环读取市下面的区
                            for (RPTAreaOrderPlanDailyEntity area : aList) {
                                if (area.getCityId().equals(city.getCityId())) {
                                    dataRow = xSheet.createRow(rowIndex++);
                                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                                    int aColumnIndex = 0;

                                    ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                                    ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

                                    dataCell = ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                    dataCell.setCellValue(null == area.getAreaName() ? "" : area.getAreaName());

                                    Class areaClass = area.getClass();

                                    aColumnIndex = writeDailyPlanOrders(days, xStyle, dataRow, area, aColumnIndex, areaClass);

                                    Double aTotalCount = StringUtils.toDouble(area.getTotal());
                                    String strATotalCount = String.format("%.0f", aTotalCount);

                                    ExportExcel.createCell(dataRow, aColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strATotalCount);
                                }

                            }// 循环读取市下面的区

                        }

                    }// 循环读取省下的市


                }// 循环读取所有的省


                //读取总计
                dataRow = xSheet.createRow(rowIndex);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                int columnIndex = 3;

                xSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
                dataCell = ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                dataCell.setCellValue(null == sumAOPD.getProvinceName() ? "" : sumAOPD.getProvinceName());


                Class totalClass = sumAOPD.getClass();

                columnIndex = writeDailyPlanOrders(days, xStyle, dataRow, sumAOPD, columnIndex, totalClass);

                Double totalCount = StringUtils.toDouble(sumAOPD.getTotal());
                String strTotalCount = String.format("%.0f", totalCount);

                ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strTotalCount);
            }

        } catch (Exception e) {
            log.error("【AreaOrderPlanDailyRptService.areaOrderPlanDayRptExport】省市区每日下单报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

    /**
     * 写入每日订单
     *
     * @return
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private int writeDailyPlanOrders(int days, Map<String, CellStyle> xStyle, Row dataRow, RPTAreaOrderPlanDailyEntity entity, int pColumnIndex, Class provinceClass) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        for (int dayIndex = 1; dayIndex <= days; dayIndex++) {
            String strGetMethodName = "getD" + dayIndex;

            Method method = provinceClass.getMethod(strGetMethodName);
            Object objGetD = method.invoke(entity);
            Double d = StringUtils.toDouble(objGetD);
            String strD = String.format("%.0f", d);

            ExportExcel.createCell(dataRow, pColumnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, strD);
        }
        return pColumnIndex;
    }
}
