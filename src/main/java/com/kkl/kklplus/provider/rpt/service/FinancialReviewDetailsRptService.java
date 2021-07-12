package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTCustomerFrozenDailyEntity;
import com.kkl.kklplus.entity.rpt.RPTFinancialReviewDetailsEntity;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.RPTSearchCondtion;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerOrderPlanDailySearch;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.FinancialReviewDetailsRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
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

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class FinancialReviewDetailsRptService extends RptBaseService {

    @Autowired
    private FinancialReviewDetailsRptMapper financialReviewDetailsRptMapper;

    public Page<RPTFinancialReviewDetailsEntity> getFinancialReviewList(RPTCustomerOrderPlanDailySearch search) {

        if (search.getPageNo() != null && search.getPageSize() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
        }
        Page<RPTFinancialReviewDetailsEntity>  page =  financialReviewDetailsRptMapper.getFinancialReviewDetailsList(search);

        Set<Long> keFuSet = new HashSet<>();
        for(RPTFinancialReviewDetailsEntity item : page){
            keFuSet.add(item.getKeFuId());
        }

        Map<String, RPTDict> abnormalMap = MSDictUtils.getDictMap("fi_charge_audit_type");//切换为微服务

        Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(keFuSet));

        RPTDict rptDictSubType;
        for(RPTFinancialReviewDetailsEntity  item : page){
            item.setCreateDate(new Date(item.getCreateDt()));

            if (namesByUserMap != null && namesByUserMap.get(item.getKeFuId()) != null) {
                item.setKeFuName(namesByUserMap.get(item.getKeFuId()));
            }

            rptDictSubType = abnormalMap.get(String.valueOf(item.getSubType()));
            if(rptDictSubType != null && StringUtils.isNotBlank(rptDictSubType.getLabel())){
                item.setAbnormalName(rptDictSubType.getLabel());
            }

        }
        return  page;
    }


    public List<RPTFinancialReviewDetailsEntity> getFinancialReviewByList(RPTCustomerOrderPlanDailySearch search) {


        List<RPTFinancialReviewDetailsEntity>  list =  financialReviewDetailsRptMapper.getFinancialReviewDetailsByList(search);

        Set<Long> keFuSet = new HashSet<>();
        for(RPTFinancialReviewDetailsEntity item : list){
            keFuSet.add(item.getKeFuId());
        }

        Map<String, RPTDict> abnormalMap = MSDictUtils.getDictMap("fi_charge_audit_type");//切换为微服务

        Map<Long, String> namesByUserMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(keFuSet));

        RPTDict rptDictSubType;
        for(RPTFinancialReviewDetailsEntity  item : list){
            item.setCreateDate(new Date(item.getCreateDt()));

            if (namesByUserMap != null && namesByUserMap.get(item.getKeFuId()) != null) {
                item.setKeFuName(namesByUserMap.get(item.getKeFuId()));
            }

            rptDictSubType = abnormalMap.get(String.valueOf(item.getSubType()));
            if(rptDictSubType != null && StringUtils.isNotBlank(rptDictSubType.getLabel())){
                item.setAbnormalName(rptDictSubType.getLabel());
            }

        }
        return  list;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
        if (searchCondition.getStartDate() != null && searchCondition.getEndDate() != null) {
            Integer rowCount = financialReviewDetailsRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }


    public SXSSFWorkbook getFinancialReviewDetailsRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerOrderPlanDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerOrderPlanDailySearch.class);
            List<RPTFinancialReviewDetailsEntity> list = getFinancialReviewByList(searchCondition);
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

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "创建时间");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单单号");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "异常类型");
            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "异常描述");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTFinancialReviewDetailsEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getKeFuName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAbnormalName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getReason());

                }

            }


        } catch (Exception e) {
            log.error("【FinancialReviewDetailsRptService.getFinancialReviewDetailsRptExport】财务审单明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
