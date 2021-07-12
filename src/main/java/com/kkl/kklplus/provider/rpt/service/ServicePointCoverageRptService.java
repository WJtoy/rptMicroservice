package com.kkl.kklplus.provider.rpt.service;


import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTServicePointCoverageEntity;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.ObjectUtils;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointCoverageRptService extends RptBaseService{

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSAreaService msAreaService;

    @Autowired
    private AreaCacheService areaCacheService;
    /**
     * 获取安维覆盖的区域
     *
     * @return
     */
    public List<RPTServicePointCoverageEntity> getServicePointCoverAreasRptData() {
        List<Long> allAreaIds = msServicePointService.findListWithAreaIds();
        Map<Long,RPTArea> provinceAreaMap = areaCacheService.getAllProvinceMap();
        List<RPTArea> cacheProvinceAreaList = provinceAreaMap.values().stream().distinct().collect(Collectors.toList());

        Map<Long,RPTArea> cityAreaMap = areaCacheService.getAllCityMap();
        List<RPTArea> cacheCityAreaList = cityAreaMap.values().stream().distinct().collect(Collectors.toList());

        Map<Long,RPTArea> countyAreaMap = areaCacheService.getAllCountyMap();
        List<RPTArea> cacheCountyAreaList = countyAreaMap.values().stream().distinct().collect(Collectors.toList());

        List<RPTArea> provinceAreaList = cacheProvinceAreaList.stream().filter(x -> allAreaIds.contains(x.getId())).collect(Collectors.toList());
        List<RPTArea> cityAreaList = cacheCityAreaList.stream().filter(x -> allAreaIds.contains(x.getId())).collect(Collectors.toList());
        List<RPTArea> countyAreaList = cacheCountyAreaList.stream().filter(x -> allAreaIds.contains(x.getId())).collect(Collectors.toList());

        List<RPTArea> townAreaList = Lists.newArrayList();

        List<Long> townAreaIds = msServicePointService.findCoverAreaList();
        if (!ObjectUtils.isEmpty(townAreaIds)) {
            Map<Long,RPTArea> cacheTownAreaMap = areaCacheService.getAllTownMap();
            List<RPTArea> cacheTownAreaList = cacheTownAreaMap.values().stream().distinct().collect(Collectors.toList());
            if (!ObjectUtils.isEmpty(cacheTownAreaList)) {
                townAreaList = cacheTownAreaList.stream().filter(x -> townAreaIds.contains(x.getId())).collect(Collectors.toList());

            }
        }

        Map<Long, List<RPTArea>> townMap = Maps.newHashMap();

        for(RPTArea area:townAreaList){
            List<RPTArea> temp = null;
            if (townMap.containsKey(area.getParent().getId())) {
                temp = townMap.get(area.getParent().getId());
            } else {
                temp = Lists.newArrayList();
                townMap.put(area.getParent().getId(), temp);
            }
            temp.add(area);
        }

        List<RPTServicePointCoverageEntity> countyRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity countyRptEntity;
        for(RPTArea area:countyAreaList){
            countyRptEntity = new RPTServicePointCoverageEntity();
            countyRptEntity.setCountyId(area.getId());
            countyRptEntity.setCountyName(area.getName());
            countyRptEntity.setParentId(area.getParent().getId());
            List<RPTArea> temp = townMap.get(area.getId());
            countyRptEntity.setAreas(temp !=null ? temp : Collections.EMPTY_LIST);
            countyRptEntityList.add(countyRptEntity);
        }

        Map<Long,List<RPTServicePointCoverageEntity>> countyMap = Maps.newHashMap();
        for(RPTServicePointCoverageEntity county:countyRptEntityList){
            List<RPTServicePointCoverageEntity> temp = null;
            if(countyMap.containsKey(county.getParentId())){
                temp = countyMap.get(county.getParentId());
            }else{
                temp = Lists.newArrayList();
                countyMap.put(county.getParentId(),temp);
            }
            temp.add(county);
        }

        List<RPTServicePointCoverageEntity> cityRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity cityRptEntity;
        for (RPTArea area : cityAreaList) {
            cityRptEntity = new RPTServicePointCoverageEntity();
            cityRptEntity.setCityId(area.getId());
            cityRptEntity.setCityName(area.getName());
            cityRptEntity.setParentId(area.getParent().getId());
            List<RPTServicePointCoverageEntity> temp = countyMap.get(area.getId());
            cityRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            cityRptEntityList.add(cityRptEntity);
        }

        Map<Long, List<RPTServicePointCoverageEntity>> cityMap = Maps.newHashMap();
        for (RPTServicePointCoverageEntity city : cityRptEntityList) {
            List<RPTServicePointCoverageEntity> temp = null;
            if (cityMap.containsKey(city.getParentId())) {
                temp = cityMap.get(city.getParentId());
            } else {
                temp = Lists.newArrayList();
                cityMap.put(city.getParentId(), temp);
            }
            temp.add(city);
        }
        List<RPTServicePointCoverageEntity> provinceRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity provinceRptEntity;
        for (RPTArea area : provinceAreaList) {
            provinceRptEntity = new RPTServicePointCoverageEntity();
            provinceRptEntity.setProvinceId(area.getId());
            provinceRptEntity.setProvinceName(area.getName());
            List<RPTServicePointCoverageEntity> temp = cityMap.get(area.getId());
            provinceRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            provinceRptEntityList.add(provinceRptEntity);
        }

        return provinceRptEntityList;
    }


    /**
     * 获取安维未覆盖的区域
     *
     * @return
     */
    public List<RPTServicePointCoverageEntity> getServicePointNoCoverAreasRptData() {

        Map<Long,RPTArea> provinceAreaMap = areaCacheService.getAllProvinceMap();
        List<RPTArea> cacheProvinceAreaList = provinceAreaMap.values().stream().distinct().collect(Collectors.toList());

        Map<Long,RPTArea> cityAreaMap = areaCacheService.getAllCityMap();
        List<RPTArea> cacheCityAreaList = cityAreaMap.values().stream().distinct().collect(Collectors.toList());

        Map<Long,RPTArea> countyAreaMap = areaCacheService.getAllCountyMap();
        List<RPTArea> cacheCountyAreaList = countyAreaMap.values().stream().distinct().collect(Collectors.toList());


        List<RPTArea> townAreaList = Lists.newArrayList();

        List<Long> townAreaIds = msServicePointService.findCoverAreaList();
        if (!ObjectUtils.isEmpty(townAreaIds)) {
            Map<Long,RPTArea> cacheTownAreaMap = areaCacheService.getAllTownMap();
            List<RPTArea> cacheTownAreaList = cacheTownAreaMap.values().stream().distinct().collect(Collectors.toList());
            if (!ObjectUtils.isEmpty(cacheTownAreaList)) {
                townAreaList = cacheTownAreaList.stream().filter(x -> !townAreaIds.contains(x.getId())).collect(Collectors.toList());

            }
        }

        Map<Long, List<RPTArea>> townMap = Maps.newHashMap();

        for(RPTArea area:townAreaList){
            List<RPTArea> temp = null;
            if (townMap.containsKey(area.getParent().getId())) {
                temp = townMap.get(area.getParent().getId());
            } else {
                temp = Lists.newArrayList();
                townMap.put(area.getParent().getId(), temp);
            }
            temp.add(area);
        }

        List<RPTServicePointCoverageEntity> countyRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity countyRptEntity;
        for(RPTArea area:cacheCountyAreaList){
            countyRptEntity = new RPTServicePointCoverageEntity();
            countyRptEntity.setCountyId(area.getId());
            countyRptEntity.setCountyName(area.getName());
            countyRptEntity.setParentId(area.getParent().getId());
            List<RPTArea> temp = townMap.get(area.getId());
            countyRptEntity.setAreas(temp !=null ? temp : Collections.EMPTY_LIST);
            countyRptEntityList.add(countyRptEntity);
        }

        Map<Long,List<RPTServicePointCoverageEntity>> countyMap = Maps.newHashMap();
        for(RPTServicePointCoverageEntity county:countyRptEntityList){
            List<RPTServicePointCoverageEntity> temp = null;
            if(countyMap.containsKey(county.getParentId())){
                temp = countyMap.get(county.getParentId());
            }else{
                temp = Lists.newArrayList();
                countyMap.put(county.getParentId(),temp);
            }
            temp.add(county);
        }

        List<RPTServicePointCoverageEntity> cityRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity cityRptEntity;
        for (RPTArea area : cacheCityAreaList) {
            cityRptEntity = new RPTServicePointCoverageEntity();
            cityRptEntity.setCityId(area.getId());
            cityRptEntity.setCityName(area.getName());
            cityRptEntity.setParentId(area.getParent().getId());
            List<RPTServicePointCoverageEntity> temp = countyMap.get(area.getId());
            cityRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            cityRptEntityList.add(cityRptEntity);
        }

        Map<Long, List<RPTServicePointCoverageEntity>> cityMap = Maps.newHashMap();
        for (RPTServicePointCoverageEntity city : cityRptEntityList) {
            List<RPTServicePointCoverageEntity> temp = null;
            if (cityMap.containsKey(city.getParentId())) {
                temp = cityMap.get(city.getParentId());
            } else {
                temp = Lists.newArrayList();
                cityMap.put(city.getParentId(), temp);
            }
            temp.add(city);
        }
        List<RPTServicePointCoverageEntity> provinceRptEntityList = Lists.newArrayList();
        RPTServicePointCoverageEntity provinceRptEntity;
        for (RPTArea area : cacheProvinceAreaList) {
            provinceRptEntity = new RPTServicePointCoverageEntity();
            provinceRptEntity.setProvinceId(area.getId());
            provinceRptEntity.setProvinceName(area.getName());
            List<RPTServicePointCoverageEntity> temp = cityMap.get(area.getId());
            provinceRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            provinceRptEntityList.add(provinceRptEntity);
        }

        return provinceRptEntityList;
    }


    public SXSSFWorkbook servicePointCoverAreasRptExport(String reportTitle) {

        SXSSFWorkbook xBook = null;
        try{
            List<RPTServicePointCoverageEntity> provinceRptEntityList = getServicePointCoverAreasRptData();
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            SXSSFSheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //=====================绘制标题行==================================
            SXSSFRow titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(30);
            ExportExcel.createCell(titleRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 2));
            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(20);

            ExportExcel.createCell(headerFirstRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 0, 0));

            ExportExcel.createCell(headerFirstRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

            ExportExcel.createCell(headerFirstRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区(县)");

            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 2, 2));

            ExportExcel.createCell(headerFirstRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "街道");
            xSheet.setColumnWidth(3,100*356);
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 3, 3));

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            //=========绘制表格数据===================
            if (provinceRptEntityList!=null){
                rowIndex++;
                for (RPTServicePointCoverageEntity entity:provinceRptEntityList) {

                    List<RPTServicePointCoverageEntity> cityList = entity.getAreaList();
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(20);
                    ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());


                    if (cityList!=null&&cityList.size()>0){
                        RPTServicePointCoverageEntity cityRptEntity = cityList.get(0);
                        ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                        RPTServicePointCoverageEntity townRptEntity = cityRptEntity.getAreaList().get(0);
                        ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getCountyName());
                        ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getCountyListString());
                        if(cityRptEntity.getAreaList() !=null && cityRptEntity.getAreaList().size()>0){
                            for(int i=1;i<cityRptEntity.getAreaList().size();i++){
                                dataRow = xSheet.createRow(rowIndex++);
                                RPTServicePointCoverageEntity town = cityRptEntity.getAreaList().get(i);
                                ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                                ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getCountyName());
                                ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getCountyListString());
                            }
                        }
                        for(int j=1;j< cityList.size();j++){
                            RPTServicePointCoverageEntity  city = cityList.get(j);
                            List<RPTServicePointCoverageEntity> countyList =city.getAreaList();
                            if(countyList !=null && countyList.size()>0){
                                for (RPTServicePointCoverageEntity aCountyList : countyList) {
                                    dataRow = xSheet.createRow(rowIndex++);
                                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCityName());
                                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getCountyName());
                                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getCountyListString());
                                }
                            }
                        }
                    }

                }
            }

        } catch (Exception e) {
            log.error("【ServicePointCoverageRptService.servicePointCoverAreasRptExport】网点覆盖报表导入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
    public SXSSFWorkbook servicePointNoCoverAreasRptExport( String reportTitle) {

        SXSSFWorkbook xBook = null;
        try{
            List<RPTServicePointCoverageEntity> provinceRptEntityList = getServicePointNoCoverAreasRptData();
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            SXSSFSheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //=====================绘制标题行==================================
            SXSSFRow titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(30);
            ExportExcel.createCell(titleRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 2));
            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(20);

            ExportExcel.createCell(headerFirstRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 0, 0));

            ExportExcel.createCell(headerFirstRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 1, 1));

            ExportExcel.createCell(headerFirstRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区(县)");

            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 2, 2));

            ExportExcel.createCell(headerFirstRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "街道");
            xSheet.setColumnWidth(3,100*356);
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 3, 3));

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            //=========绘制表格数据===================
            if (provinceRptEntityList!=null){
                rowIndex++;
                for (RPTServicePointCoverageEntity entity:provinceRptEntityList) {

                    List<RPTServicePointCoverageEntity> cityList = entity.getAreaList();
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(20);
                    ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());


                    if (cityList!=null&&cityList.size()>0){
                        RPTServicePointCoverageEntity cityRptEntity = cityList.get(0);
                        ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                        RPTServicePointCoverageEntity townRptEntity = cityRptEntity.getAreaList().get(0);
                        ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getCountyName());
                        ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getCountyListString());
                        if(cityRptEntity.getAreaList() !=null && cityRptEntity.getAreaList().size()>0){
                            for(int i=1;i<cityRptEntity.getAreaList().size();i++){
                                dataRow = xSheet.createRow(rowIndex++);
                                RPTServicePointCoverageEntity town = cityRptEntity.getAreaList().get(i);
                                ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                                ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getCountyName());
                                ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getCountyListString());
                            }
                        }
                        for(int j=1;j< cityList.size();j++){
                            RPTServicePointCoverageEntity  city = cityList.get(j);
                            List<RPTServicePointCoverageEntity> countyList =city.getAreaList();
                            if(countyList !=null && countyList.size()>0){
                                for (RPTServicePointCoverageEntity aCountyList : countyList) {
                                    dataRow = xSheet.createRow(rowIndex++);
                                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCityName());
                                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getCountyName());
                                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getCountyListString());
                                }
                            }
                        }
                    }

                }
            }

        } catch (Exception e) {
            log.error("【ServicePointCoverageRptService.servicePointNoCoverAreasRptExport】网点无覆盖报表导入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
