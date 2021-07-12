package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCrushCoverageEntity;
import com.kkl.kklplus.entity.rpt.search.RPTGradedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
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
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class TravelCoverageRptService extends RptBaseService {

    @Autowired
    private MSAreaService msAreaService;

    @Autowired
    private AreaCacheService areaCacheService;



    /**
     * 远程覆盖区域
     *
     * @return
     */
    public List<RPTCrushCoverageEntity> getTravelCoverAreasRptData(RPTGradedOrderSearch rptGradedOrderSearch) {

        Map<Long,RPTArea> cacheProvinceAreaList = areaCacheService.getAllProvinceMap();
        Map<Long,RPTArea> cacheCityAreaList = areaCacheService.getAllCityMap();
        Map<Long,RPTArea> cacheCountyAreaList = areaCacheService.getAllCountyMap();


        List<RPTArea> townAreaList = Lists.newArrayList();
        List<RPTArea> noTownAreaList = Lists.newArrayList();

        List<Long> townAreaIds = MSAreaUtils.getTravelAreaMap(rptGradedOrderSearch.getProductCategoryIds());
        if (!ObjectUtils.isEmpty(townAreaIds)) {
            Map<Long,RPTArea> cacheTownAreaMap =  areaCacheService.getAllTownMap();
            if (!ObjectUtils.isEmpty(cacheTownAreaMap)) {
                List<RPTArea> cacheTownAreaList = cacheTownAreaMap.entrySet().stream().map(i -> i.getValue()).distinct().collect(Collectors.toList());
                townAreaList = cacheTownAreaList.stream().filter(x -> townAreaIds.contains(x.getId())).collect(Collectors.toList());
                noTownAreaList = cacheTownAreaList.stream().filter(x -> !townAreaIds.contains(x.getId())).collect(Collectors.toList());

            }
        }else {
            Map<Long,RPTArea> cacheTownAreaMap =  areaCacheService.getAllTownMap();
            noTownAreaList = cacheTownAreaMap.entrySet().stream().map(i -> i.getValue()).distinct().collect(Collectors.toList());
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

        Map<Long, List<RPTArea>> noTownMap = Maps.newHashMap();

        for(RPTArea area:noTownAreaList){
            List<RPTArea> temp = null;
            if (noTownMap.containsKey(area.getParent().getId())) {
                temp = noTownMap.get(area.getParent().getId());
            } else {
                temp = Lists.newArrayList();
                noTownMap.put(area.getParent().getId(), temp);
            }
            temp.add(area);
        }

        String areaName;
        String noareaName;
        List<RPTCrushCoverageEntity> countyRptEntityList = Lists.newArrayList();
        RPTCrushCoverageEntity countyRptEntity;
        for(RPTArea area:cacheCountyAreaList.values()){
            countyRptEntity = new RPTCrushCoverageEntity();
            countyRptEntity.setCountyId(area.getId());
            countyRptEntity.setCountyName(area.getName());
            countyRptEntity.setParentId(area.getParent().getId());
            List<RPTArea> temp = townMap.get(area.getId());
            List<RPTArea> noTemp = noTownMap.get(area.getId());
            if(temp != null && temp.size()>0){
                areaName = getCountyListString(temp);
                countyRptEntity.setAreaName(areaName);
            }
            if(noTemp != null && noTemp.size()>0){
                noareaName = getCountyListString(noTemp);
                countyRptEntity.setNoareaName(noareaName);
            }
            countyRptEntityList.add(countyRptEntity);
        }

        Map<Long,List<RPTCrushCoverageEntity>> countyMap = Maps.newHashMap();
        for(RPTCrushCoverageEntity county:countyRptEntityList){
            List<RPTCrushCoverageEntity> temp = null;
            if(countyMap.containsKey(county.getParentId())){
                temp = countyMap.get(county.getParentId());
            }else{
                temp = Lists.newArrayList();
                countyMap.put(county.getParentId(),temp);
            }
            temp.add(county);
        }

        List<RPTCrushCoverageEntity> cityRptEntityList = Lists.newArrayList();
        RPTCrushCoverageEntity cityRptEntity;
        for (RPTArea area : cacheCityAreaList.values()) {
            cityRptEntity = new RPTCrushCoverageEntity();
            cityRptEntity.setCityId(area.getId());
            cityRptEntity.setCityName(area.getName());
            cityRptEntity.setParentId(area.getParent().getId());
            List<RPTCrushCoverageEntity> temp = countyMap.get(area.getId());
            cityRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            cityRptEntityList.add(cityRptEntity);
        }

        Map<Long, List<RPTCrushCoverageEntity>> cityMap = Maps.newHashMap();
        for (RPTCrushCoverageEntity city : cityRptEntityList) {
            List<RPTCrushCoverageEntity> temp = null;
            if (cityMap.containsKey(city.getParentId())) {
                temp = cityMap.get(city.getParentId());
            } else {
                temp = Lists.newArrayList();
                cityMap.put(city.getParentId(), temp);
            }
            temp.add(city);
        }
        List<RPTCrushCoverageEntity> provinceRptEntityList = Lists.newArrayList();
        RPTCrushCoverageEntity provinceRptEntity;
        for (RPTArea area : cacheProvinceAreaList.values()) {
            provinceRptEntity = new RPTCrushCoverageEntity();
            provinceRptEntity.setProvinceId(area.getId());
            provinceRptEntity.setProvinceName(area.getName());
            List<RPTCrushCoverageEntity> temp = cityMap.get(area.getId());
            provinceRptEntity.setAreaList(temp != null ? temp : Collections.EMPTY_LIST);
            provinceRptEntityList.add(provinceRptEntity);
        }

        return provinceRptEntityList;
    }

    public String getCountyListString(List<RPTArea> temp) {
        StringBuilder stringBuilder = new StringBuilder();
        RPTArea rptArea;
        for(int i = 0; i<=temp.size()-1 ; i++){
            rptArea  = temp.get(i);
            if(i <temp.size()-1){
                stringBuilder.append(rptArea.getName() + ",");
            }else {
                stringBuilder.append(rptArea.getName());
            }

        }
        return stringBuilder.toString();
    }


    public SXSSFWorkbook travelCoverAreasRptExport(String searchConditionJson,String reportTitle) {

        SXSSFWorkbook xBook = null;
        try{
            RPTGradedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTGradedOrderSearch.class);
            List<RPTCrushCoverageEntity> provinceRptEntityList = getTravelCoverAreasRptData(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0,4));
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

            ExportExcel.createCell(headerFirstRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "覆盖区域");
            xSheet.setColumnWidth(3,100*356);
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 3, 3));

            ExportExcel.createCell(headerFirstRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未覆盖区域");
            xSheet.setColumnWidth(4,100*356);
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum()+1, 4, 4));

            xSheet.createFreezePane(0, rowIndex+1); // 冻结单元格(x, y)
            //=========绘制表格数据===================
            if (provinceRptEntityList!=null){
                rowIndex++;
                for (RPTCrushCoverageEntity entity:provinceRptEntityList) {

                    List<RPTCrushCoverageEntity> cityList = entity.getAreaList();
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(20);
                    ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());


                    if (cityList!=null&&cityList.size()>0){
                        RPTCrushCoverageEntity cityRptEntity = cityList.get(0);
                        ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                        RPTCrushCoverageEntity townRptEntity = cityRptEntity.getAreaList().get(0);
                        ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getCountyName());
                        ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getAreaName());
                        ExportExcel.createCell(dataRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, townRptEntity.getNoareaName());
                        if(cityRptEntity.getAreaList() !=null && cityRptEntity.getAreaList().size()>0){
                            for(int i=1;i<cityRptEntity.getAreaList().size();i++){
                                dataRow = xSheet.createRow(rowIndex++);
                                RPTCrushCoverageEntity town = cityRptEntity.getAreaList().get(i);
                                ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, cityRptEntity.getCityName());
                                ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getCountyName());
                                ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getAreaName());
                                ExportExcel.createCell(dataRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, town.getNoareaName());
                            }
                        }
                        for(int j=1;j< cityList.size();j++){
                            RPTCrushCoverageEntity  city = cityList.get(j);
                            List<RPTCrushCoverageEntity> countyList =city.getAreaList();
                            if(countyList !=null && countyList.size()>0){
                                for (RPTCrushCoverageEntity aCountyList : countyList) {
                                    dataRow = xSheet.createRow(rowIndex++);
                                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, city.getCityName());
                                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getCountyName());
                                    ExportExcel.createCell(dataRow,3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getAreaName());
                                    ExportExcel.createCell(dataRow,4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, aCountyList.getNoareaName());
                                }
                            }
                        }
                    }

                }
            }

        } catch (Exception e) {
            log.error("【TravelCoverageRptService.travelCoverAreasRptExport】远程区域报表导入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
