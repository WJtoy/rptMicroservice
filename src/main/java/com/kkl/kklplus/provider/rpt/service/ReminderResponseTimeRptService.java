package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainEntity;
import com.kkl.kklplus.entity.rpt.RPTReminderResponseTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTReminderResponseTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.ReminderResponseTimeRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
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

import javax.annotation.Resource;
import java.util.*;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ReminderResponseTimeRptService extends RptBaseService {

    @Resource
    private ReminderResponseTimeRptMapper reminderResponseTimeRptMapper;

    @Autowired
    private AreaCacheService areaCacheService;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    public Page<RPTReminderResponseTimeEntity> getReminderResponseTimeRptData(RPTReminderResponseTimeSearch search) {
        Page<RPTReminderResponseTimeEntity> returnPage = new Page<>();
        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            returnPage = reminderResponseTimeRptMapper.getReminderResponseTimerListByPaging(search);
            Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
            Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
            Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
            Set<Long> engineerSet = new HashSet<>();
            Set<Long> keFuSet = new HashSet<>();
            Set<Long> servicePointSet = new HashSet<>();
            for(RPTReminderResponseTimeEntity item : returnPage){
                engineerSet.add(item.getEngineerId());
                keFuSet.add(item.getKeFuId());
                if(item.getServicePointId()!=null && item.getServicePointId()!=0){
                    servicePointSet.add(item.getServicePointId());
                }
            }
            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerSet), Arrays.asList("id", "name"));

            Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointSet), Arrays.asList("id","name"), null);

            Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(keFuSet));

            RPTEngineer engineer;
            MDServicePointViewModel ServicePointViewModel;
            for (RPTReminderResponseTimeEntity entity : returnPage) {

                engineer = engineerMap.get(entity.getEngineerId());
                entity.setEngineerName(engineer == null ? "" : engineer.getName());

                if(servicePointMap != null){
                    ServicePointViewModel = servicePointMap.get(entity.getServicePointId());
                    entity.setServicePointName(ServicePointViewModel == null ? "" : ServicePointViewModel.getName());
                }
                if(namesByUserMap != null){
                    String keFuName =  namesByUserMap.get(entity.getKeFuId());
                    if(StringUtils.isNotBlank(keFuName)){
                        entity.setKeFuName(keFuName);
                    }
                }

                if(provinceMap != null){
                    entity.setProvinceName(provinceMap.get(entity.getProvinceId()).getName());
                }
                if(cityMap != null) {
                    entity.setCityName(cityMap.get(entity.getCityId()).getName());
                }
                if(areaMap != null){
                    entity.setAreaName(areaMap.get(entity.getAreaId()).getName());
                }

                entity.setCreateDate(new Date(entity.getCreateDt()));
                if(entity.getProcessDt() != 0){
                    entity.setProcessDate(new Date(entity.getProcessDt()));
                }
                if(entity.getCompleteDt() != null && entity.getCompleteDt() != 0){
                    entity.setCompleteDate(new Date(entity.getCompleteDt()));
                }

            }
        }
        return returnPage;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTReminderResponseTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTReminderResponseTimeSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getBeginDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = reminderResponseTimeRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook reminderResponseTimeExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTReminderResponseTimeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTReminderResponseTimeSearch.class);
            Page<RPTReminderResponseTimeEntity> list = getReminderResponseTimeRptData(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(5000);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;


            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 19));


            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单号");
            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");
            ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            ExportExcel.createCell(headerFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            ExportExcel.createCell(headerFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");
            ExportExcel.createCell(headerFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "催单单号");
            ExportExcel.createCell(headerFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "催单次数");
            ExportExcel.createCell(headerFirstRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "发起时间");
            ExportExcel.createCell(headerFirstRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "催单意见");
            ExportExcel.createCell(headerFirstRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "催单回复人");
            ExportExcel.createCell(headerFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "回复时间");
            ExportExcel.createCell(headerFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "回复用时");
            ExportExcel.createCell(headerFirstRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "回复结果");
            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "师傅");
            ExportExcel.createCell(headerFirstRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评时间");
            ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客评用时");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            //====================================================绘制表格数据单元格============================================================
            int totalQty = 0;
            double totalProcessTimeliness = 0.0;
            double totalOrderTimeliness = 0.0;


            if (list != null) {
                for (int i = 0; i < list.size(); i++) {
                    RPTReminderResponseTimeEntity entity = list.get(i);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderNo());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getKeFuName());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getServicePointName());
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getUserName());
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getUserPhone());
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProvinceName());
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getCityName());
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getAreaName());
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderNo());
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderTimes());
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(entity.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getReminderRemark());
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProcessBy());
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(entity.getProcessDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProcessTimeliness());
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getProcessRemark());
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getEngineerName());
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(entity.getCompleteDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderTimeliness());
                    totalQty = totalQty + entity.getReminderTimes();
                    totalProcessTimeliness = totalProcessTimeliness + entity.getProcessTimeliness();
                    totalOrderTimeliness = totalOrderTimeliness + entity.getOrderTimeliness();
                }
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 9));

                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);

                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 11, 14));

                ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalProcessTimeliness);
                ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 16, 18));
                ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOrderTimeliness);
            }
        } catch (Exception e) {
            log.error("【ReminderResponseTimeRptService.reminderResponseTimeExport】催单回复时效报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }

        return xBook;
    }
}
