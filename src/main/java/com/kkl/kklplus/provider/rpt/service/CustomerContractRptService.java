package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.b2bcenter.md.B2BSign;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.praise.PraiseStatusEnum;
import com.kkl.kklplus.entity.rpt.CustomerSignTypeEnum;
import com.kkl.kklplus.entity.rpt.RPTCustomerFinanceEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerSignSearch;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.CustomerPraiseDetailsRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.service.B2BCustomerMappingService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
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

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerContractRptService extends RptBaseService {

    @Autowired
    private B2BCustomerMappingService b2BCustomerMappingService;

    public MSPage<B2BSign> getCustomerSignDetailsList(RPTCustomerSignSearch search) {
        MSPage<B2BSign> page = new MSPage<>();

        B2BSign b2BSign = new B2BSign();
        if (search.getMallId() != null && search.getMallId() != 0) {
            b2BSign.setMallId(search.getMallId());
        }
        if (StringUtils.isNotBlank(search.getMallName())) {
            b2BSign.setMallName(search.getMallName());
        }

        if (StringUtils.isNotBlank(search.getMobile())) {
            b2BSign.setMobile(search.getMobile());
        }

        if (search.getBeginDate() != null && search.getBeginDate() != 0) {
            b2BSign.setBeginApplyTime(search.getBeginDate());
        }

        if (search.getEndDate() != null && search.getEndDate() != 0) {
            b2BSign.setEndApplyTime(search.getEndDate());
        }

        if(search.getStatus()!= null){
            b2BSign.setSignStatus(search.getStatus());
        }

        page.setPageNo(search.getPageNo());
        page.setPageSize(search.getPageSize());


        page = b2BCustomerMappingService.findCustomerContract(page,b2BSign);

        if(page != null && page.getList()!=null){
            for(B2BSign b2BSign1 : page.getList()){
                if(b2BSign1.getSignStatus() != null ){
                    b2BSign1.setAttributes(String.valueOf(CustomerSignTypeEnum.fromCode(Integer.valueOf(b2BSign1.getSignStatus())).msg));
                }

                b2BSign1.setApplyDate(new Date(b2BSign1.getApplyTime()));
                b2BSign1.setDataSource(16);

            }
        }

        return page;
    }


    public List<B2BSign> getCustomerSignDetailsExport(RPTCustomerSignSearch search) {
        B2BSign b2BSign = new B2BSign();

        if (search.getMallId() != null && search.getMallId() != 0) {
            b2BSign.setMallId(search.getMallId());
        }
        if (StringUtils.isNotBlank(search.getMallName())) {
            b2BSign.setMallName(search.getMallName());
        }

        if (StringUtils.isNotBlank(search.getMobile())) {
            b2BSign.setMobile(search.getMobile());
        }

        if (search.getBeginDate() != null && search.getBeginDate() != 0) {
            b2BSign.setBeginApplyTime(search.getBeginDate());
        }

        if (search.getEndDate() != null && search.getEndDate() != 0) {
            b2BSign.setEndApplyTime(search.getEndDate());
        }

        if(search.getStatus()!= null){
            b2BSign.setSignStatus(search.getStatus());
        }

        List<B2BSign>  b2BSignList= b2BCustomerMappingService.findCustomerContractList(b2BSign);
        Map<String, RPTDict>  dataSourceMap = MSDictUtils.getDictMap("order_data_source");
            for(B2BSign b2BSign1 : b2BSignList){
                if(b2BSign1.getSignStatus() != null ){
                    b2BSign1.setAttributes(String.valueOf(CustomerSignTypeEnum.fromCode(Integer.valueOf(b2BSign1.getSignStatus())).msg));
                }

                String key = String.valueOf(b2BSign1.getDataSource());

                if (dataSourceMap.get(key) != null && dataSourceMap.get(key).getLabel() != null) {
                    b2BSign1.setDataType(dataSourceMap.get(key).getLabel());
                }
                b2BSign1.setApplyDate(new Date(b2BSign1.getApplyTime()));
                b2BSign1.setDataSource(16);

            }

        return b2BSignList;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        B2BSign b2BSign = new B2BSign();
        RPTCustomerSignSearch search = redisGsonService.fromJson(searchConditionJson, RPTCustomerSignSearch.class);
        if (search.getBeginDate() != null && search.getEndDate() != null) {

            if (search.getMallId() != null && search.getMallId() != 0) {
                b2BSign.setMallId(search.getMallId());
            }
            if (StringUtils.isNotBlank(search.getMallName())) {
                b2BSign.setMallName(search.getMallName());
            }

            if (StringUtils.isNotBlank(search.getMobile())) {
                b2BSign.setMobile(search.getMobile());
            }

            if (search.getBeginDate() != null && search.getBeginDate() != 0) {
                b2BSign.setBeginApplyTime(search.getBeginDate());
            }

            if (search.getEndDate() != null && search.getEndDate() != 0) {
                b2BSign.setEndApplyTime(search.getEndDate());
            }

            if(search.getStatus()!= null){
                b2BSign.setSignStatus(search.getStatus());
            }
            List<B2BSign>  b2BSignList= b2BCustomerMappingService.findCustomerContractList(b2BSign);
            if(b2BSignList == null || b2BSignList.size() == 0){
                result = false;
            }
        }
        return result;
    }


    public SXSSFWorkbook customerSignDetailsRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerSignSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerSignSearch.class);
            List<B2BSign> list = getCustomerSignDetailsExport(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "数据源");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约单号");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺ID");

            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务名称");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "申请时间");

            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系人");
            ExportExcel.createCell(headRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系电话");

            ExportExcel.createCell(headRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "座机");
            ExportExcel.createCell(headRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    B2BSign rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getDataType());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSignOrderSn());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAttributes());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getMallId());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getMallName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServType());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getApplyDate(), "yyyy-MM-dd HH:mm:ss"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getContactName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getMobile());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getTelephone());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRemarks());

                }

            }


        } catch (Exception e) {
            log.error("【CustomerContractRptService.customerSignDetailsRptExport】客服好评明细写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
