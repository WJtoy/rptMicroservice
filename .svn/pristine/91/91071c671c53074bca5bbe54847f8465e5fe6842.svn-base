package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.RPTCrushCoverageEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuAreaEntity;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.KeFuAreaEntity;
import com.kkl.kklplus.provider.rpt.mapper.KeFuAreaRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
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
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class KeFuAreaRptService extends RptBaseService {

    @Autowired
    private KeFuAreaRptMapper keFuAreaRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;


    /**
     * 客服区域
     *
     * @return
     */
    public List<RPTKeFuAreaEntity> getKeFuAreasRptData() {

        int systemId = RptCommonUtils.getSystemId();
        List<RPTKeFuAreaEntity>   keFuAreaList = Lists.newArrayList();
//        List<RPTKeFuAreaEntity>   keFuAreaLists = keFuAreaRptMapper.getKeFuAreaList(systemId);
        List<KeFuAreaEntity>   keFuLists = Lists.newArrayList();
        List<KeFuAreaEntity>   keFuProvinceLists = keFuAreaRptMapper.getKeFuProvinceList(systemId);
        List<KeFuAreaEntity>   keFuCityLists = keFuAreaRptMapper.getKeFuCityList(systemId);
        List<KeFuAreaEntity>   keFuAreasLists = keFuAreaRptMapper.getKeFuAreasList(systemId);
        List<KeFuAreaEntity>   keFuCountryLists = keFuAreaRptMapper.getKeFuCountryList(systemId);
        keFuLists.addAll(keFuProvinceLists);
        keFuLists.addAll(keFuCityLists);
        keFuLists.addAll(keFuAreasLists);
        keFuLists.addAll(keFuCountryLists);


        Map<Long, List<KeFuAreaEntity>> keFuAreaEntityMap = keFuLists.stream().collect(Collectors.groupingBy(KeFuAreaEntity::getUserId));


        Map<Long, RPTUser> keFuMap = MSUserUtils.getMapByUserType(2);
        RPTUser rptUser;
        RPTKeFuAreaEntity rptKeFuAreaEntity;
        Map<Long, RPTArea> countyMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        RPTArea rptArea;
        for(List<KeFuAreaEntity> item : keFuAreaEntityMap.values()){
            rptUser = keFuMap.get(item.get(0).getUserId());
            rptKeFuAreaEntity = new RPTKeFuAreaEntity();
            StringBuilder sb = new StringBuilder();
            if(rptUser != null){
                int index = 0;
                for(KeFuAreaEntity keFuAreaEntity : item){
                    if(keFuAreaEntity.getAreaType() == 0){
                        sb.append("中国");
                        index++;
                    }
                    if(keFuAreaEntity.getAreaType() == 2){
                        rptArea = provinceMap.get(keFuAreaEntity.getProvinceId());
                        if(index != 0 && rptArea != null){
                            sb.append(",");
                        }
                        if(rptArea != null && rptArea.getFullName() != null){
                            sb.append(rptArea.getFullName());
                            index++;
                        }
                    }
                    if(keFuAreaEntity.getAreaType() == 3){
                        rptArea = cityMap.get(keFuAreaEntity.getCityId());
                        if(index != 0 && rptArea != null){
                            sb.append(",");
                        }
                        if(rptArea != null && rptArea.getFullName() != null){
                            sb.append(rptArea.getFullName());
                            index++;
                        }

                    }
                    if(keFuAreaEntity.getAreaType() == 4){
                        rptArea = countyMap.get(keFuAreaEntity.getAreaId());
                        if(index != 0 && rptArea != null){
                            sb.append(",");
                        }
                        if(rptArea != null && rptArea.getFullName() != null){
                            sb.append(rptArea.getFullName());
                            index++;
                        }
                    }
                }
                rptKeFuAreaEntity.setAreaName(sb.toString());
                rptKeFuAreaEntity.setKefuName(rptUser.getName());
                rptKeFuAreaEntity.setQq(rptUser.getQq());
                keFuAreaList.add(rptKeFuAreaEntity);
            }
        }

        return keFuAreaList;
    }


    public SXSSFWorkbook keFuAreasRptExport(String reportTitle) {

        SXSSFWorkbook xBook = null;
        try{
            List<RPTKeFuAreaEntity> provinceRptEntityList = getKeFuAreasRptData();
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            SXSSFSheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //=====================绘制标题行==================================
            SXSSFRow titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(30);
            ExportExcel.createCell(titleRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0,3));
            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(20);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headerFirstRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "姓名");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headerFirstRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "qq");


            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headerFirstRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区域");
            xSheet.setColumnWidth(2,100*356);

            xSheet.createFreezePane(0, rowIndex+1); // 冻结单元格(x, y)
            //=========绘制表格数据===================
            if (provinceRptEntityList!=null){
                rowIndex++;
                for (RPTKeFuAreaEntity entity:provinceRptEntityList) {
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(20);
                    ExportExcel.createCell(dataRow,0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getKefuName());
                    ExportExcel.createCell(dataRow,1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getQq());
                    ExportExcel.createCell(dataRow,2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAreaName());

                }

            }

        } catch (Exception e) {
            log.error("【KeFuAreaRptService.keFuAreasRptExport】客服区域报表导入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
