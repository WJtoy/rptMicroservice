package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.md.MDMaterial;
import com.kkl.kklplus.entity.rpt.RPTMasterApplyEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointPaySummaryEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.MasterApplyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSMaterialMasterService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
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

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MasterApplyRptService extends RptBaseService {

    @Autowired
    private MasterApplyRptMapper masterApplyRptMapper;

    @Autowired
    private MSMaterialMasterService msMaterialMasterService;

    @Autowired
    private AreaCacheService areaCacheService;

    public Page<RPTMasterApplyEntity> getMasterApplyList(RPTComplainStatisticsDailySearch search) {
        Page<RPTMasterApplyEntity> page;

        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }

        Date beginDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());

        page = masterApplyRptMapper.getMasterApplyPage(beginDate,endDate,search.getCustomerId(),search.getQuarter());
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<String, RPTDict> materialMap = MSDictUtils.getDictMap("MaterialStatus");
        RPTDict materialDict;
        RPTArea rptArea;
        Set<Long> materialSet = new HashSet<>();
        for(RPTMasterApplyEntity item : page){
            rptArea = areaMap.get(item.getAreaId());
            if (rptArea != null) {
                item.setAddress(rptArea.getFullName() + " " + item.getUserAddress());
            }
            materialSet.add(item.getMaterialId());
            if(item.getStatus() != null ){
                materialDict = materialMap.get(String.valueOf(item.getStatus()));
                if(materialDict != null && StringUtils.isNotBlank(materialDict.getLabel())){
                    item.setStatusName(materialDict.getLabel());
                }
            }
        }

        List<MDMaterial> mdMaterials   = msMaterialMasterService.findMaterialWithIds(Lists.newArrayList(materialSet));
        Map<Long, MDMaterial> materialsMap = mdMaterials.stream().collect(Collectors.toMap(MDMaterial::getId, Function.identity()));
        MDMaterial mdMaterial;
        for(RPTMasterApplyEntity item : page){
            if(materialsMap != null){
                mdMaterial = materialsMap.get(item.getMaterialId());
                if(mdMaterial!=null){
                    item.setMasterName(mdMaterial.getName());
                }
            }

        }
        return page;
    }


    public List<RPTMasterApplyEntity> getMasterApplyListData(RPTComplainStatisticsDailySearch search) {
        List<RPTMasterApplyEntity> list;

        Date beginDate = new Date(search.getStartDate());
        Date endDate = new Date(search.getEndDate());

        list = masterApplyRptMapper.getMasterApplyList(beginDate,endDate,search.getCustomerId(),search.getQuarter());
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<String, RPTDict> materialMap = MSDictUtils.getDictMap("MaterialStatus");
        RPTDict materialDict;
        RPTArea rptArea;
        Set<Long> materialSet = new HashSet<>();
        for(RPTMasterApplyEntity item : list){
            rptArea = areaMap.get(item.getAreaId());
            if (rptArea != null) {
                item.setAddress(rptArea.getFullName() + " " + item.getUserAddress());
            }
            materialSet.add(item.getMaterialId());
            if(item.getStatus() != null ){
                materialDict = materialMap.get(String.valueOf(item.getStatus()));
                if(materialDict != null && StringUtils.isNotBlank(materialDict.getLabel())){
                    item.setStatusName(materialDict.getLabel());
                }
            }
        }

        List<MDMaterial> mdMaterials   = msMaterialMasterService.findMaterialWithIds(Lists.newArrayList(materialSet));
        Map<Long, MDMaterial> materialsMap = mdMaterials.stream().collect(Collectors.toMap(MDMaterial::getId, Function.identity()));
        MDMaterial mdMaterial;
        for(RPTMasterApplyEntity item : list){
            if(materialsMap != null){
                mdMaterial = materialsMap.get(item.getMaterialId());
                if(mdMaterial!=null){
                    item.setMasterName(mdMaterial.getName());
                }
            }

        }
        return list;
    }


    /**
     * 检查开发均单费用报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
        Date beginDate = new Date(searchCondition.getStartDate());
        Date endDate = new Date(searchCondition.getEndDate());
        if (searchCondition != null) {
            Integer rowCount = masterApplyRptMapper.hasReportData(beginDate,endDate,searchCondition.getCustomerId(),searchCondition.getQuarter());
            result = rowCount > 0;
        }
        return result;
    }







    public SXSSFWorkbook MasterApplyRptExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;

        try {
            RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
            List<RPTMasterApplyEntity> list = getMasterApplyListData(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "订单号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "电话");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "地址");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 9));
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件申请");


            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 10, 10));
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成时间");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 11, 11));
            ExportExcel.createCell(headFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件申请时间");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "名称");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "数量");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "是否返件");
            ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "返件快递单号");
            ExportExcel.createCell(headSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if(list != null && list.size() > 0){

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTMasterApplyEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUserName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePhone());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAddress());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getDescription());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getMasterName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getQty());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReturnFlag() == 1 ? "是":"否");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getExpressNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getStatusName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getCloseDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,  DateUtils.formatDate(rowData.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));

                }

            }


        } catch (Exception e) {
            log.error("【MasterApplyRptService.MasterApplyRptExport】配件报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }


}
