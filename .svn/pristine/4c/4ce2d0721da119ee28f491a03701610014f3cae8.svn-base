package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.praise.PraiseStatusEnum;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointPraiseDetailsRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
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
public class ServicePointPraiseDetailsRptService extends RptBaseService {

    @Autowired
    private ServicePointPraiseDetailsRptMapper servicePointPraiseDetailsRptMapper;


    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;

    public Page<RPTKeFuPraiseDetailsEntity> getServicePointPraiseDetailsList(RPTKeFuCompleteTimeSearch search) {



        Page<RPTKeFuPraiseDetailsEntity> page;

        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        page = servicePointPraiseDetailsRptMapper.getServicePointPraiseDetailsPage(search);

        Set<Long> engineerSet = new HashSet<>();
        for(RPTKeFuPraiseDetailsEntity item : page){
            engineerSet.add(item.getEngineerId());
        }
        List<Long> engineerIds = Lists.newArrayList(engineerSet);
        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));


        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap =areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        String countyName = "";
        String cityName = "";
        String provinceName = "";
        RPTDict praiseDict;
        RPTEngineer engineer;
        for(RPTKeFuPraiseDetailsEntity item : page){
            item.setCreateDate(new Date(item.getCreateDt()));

            if(item.getStatus() != null && item.getStatus().getValue() != null){
                praiseDict = new RPTDict();
                praiseDict.setLabel(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).msg);
                praiseDict.setValue(String.valueOf(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).code));
                item.setStatus(praiseDict);
            }


            engineer = engineerMap.get(item.getEngineerId());
            item.setEngineerName(engineer == null ? "" : engineer.getName());


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


    public List<RPTKeFuPraiseDetailsEntity> getServicePointPraiseDetail(RPTKeFuCompleteTimeSearch search) {

        List<RPTKeFuPraiseDetailsEntity> list = servicePointPraiseDetailsRptMapper.getServicePointPraiseDetailsPage(search);



        Set<Long> engineerSet = new HashSet<>();
        for(RPTKeFuPraiseDetailsEntity item : list){
            engineerSet.add(item.getEngineerId());
        }
        List<Long> engineerIds = Lists.newArrayList(engineerSet);
        Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));



        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
        String countyName = "";
        String cityName = "";
        String provinceName = "";
        RPTDict praiseDict;
        RPTEngineer engineer;
        for(RPTKeFuPraiseDetailsEntity item : list){
            item.setCreateDate(new Date(item.getCreateDt()));


            if(item.getStatus() != null && item.getStatus().getValue() != null){
                praiseDict = new RPTDict();
                praiseDict.setLabel(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).msg);
                praiseDict.setValue(String.valueOf(PraiseStatusEnum.fromCode(Integer.valueOf(item.getStatus().getValue())).code));
                item.setStatus(praiseDict);
            }


            engineer = engineerMap.get(item.getEngineerId());
            item.setEngineerName(engineer == null ? "" : engineer.getName());

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
            Integer rowCount;
            Integer keFuPraiseSum = servicePointPraiseDetailsRptMapper.getServicePointPraiseSum(searchCondition);
            rowCount = keFuPraiseSum;
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook servicePointPraiseDetailsRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTKeFuCompleteTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTKeFuCompleteTimeSearch.class);
            List<RPTKeFuPraiseDetailsEntity> list = getServicePointPraiseDetail(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单单号");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "创建时间");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区域");
            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户姓名");

            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "师傅");

            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");

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

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getStatus().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAreaName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserPhone());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserAddress());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerName());


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
                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 7));
                ExportExcel.createCell(sumRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalPraise);

            }


        } catch (Exception e) {
            log.error("【KeFuPraiseDetailsRptService.keFuPraiseDetailsRptExport】客服好评明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
