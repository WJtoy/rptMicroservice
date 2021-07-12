package com.kkl.kklplus.provider.rpt.customer.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.rpt.RPTCancelledOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.common.RPTMiddleTableEnum;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.entity.LongThreeTuple;
import com.kkl.kklplus.provider.rpt.mapper.CustomerChargeSummaryRptNewMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.service.CancelledOrderRptService;
import com.kkl.kklplus.provider.rpt.service.CompletedOrderRptService;
import com.kkl.kklplus.provider.rpt.service.CustomerWriteOffRptService;
import com.kkl.kklplus.provider.rpt.service.RptBaseService;
import com.kkl.kklplus.provider.rpt.utils.*;
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
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 客户对账单 - 工单数量与消费金额汇总
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CtCustomerChargeSummaryRptNewService extends RptBaseService {

    @Resource
    private CustomerChargeSummaryRptNewMapper customerChargeSummaryRptNewMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private CtCompletedOrderRptService ctCompletedOrderRptService;

    @Autowired
    private CtCancelledOrderRptService ctCancelledOrderRptService;

    @Autowired
    private CtCustomerWriteOffRptService ctCustomerWriteOffRptService;



    //endregion 公开

    public RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryNew(RPTCustomerChargeSearch search) {
        RPTCustomerChargeSummaryMonthlyEntity result = null;
        if (search != null && search.getCustomerId() != null && search.getCustomerId() > 0
                && search.getSelectedYear() != null && search.getSelectedYear() > 0
                && search.getSelectedMonth() != null && search.getSelectedMonth() > 0) {
            int yearMonth = generateYearMonth(search.getSelectedYear(), search.getSelectedMonth());
            int currentYearMonth = generateYearMonth(new Date());

//            if (yearMonth == currentYearMonth) {
//                result = getCustomerChargeSummaryMonthlyByCurrentMonth(search.getCustomerId(), search.getSelectedYear(), search.getSelectedMonth());
//            } else
            if (yearMonth <= currentYearMonth) {
                result = getCustomerChargeSummaryMonthly(search.getCustomerId(), yearMonth);
            } else {
                result = new RPTCustomerChargeSummaryMonthlyEntity();
            }
        }
        return result;
    }


    private RPTCustomerChargeSummaryMonthlyEntity getCustomerChargeSummaryMonthly(long customerId, int yearmonth) {
        int systemId = RptCommonUtils.getSystemId();
        RPTCustomerChargeSummaryMonthlyEntity orderQtyMonthly = customerChargeSummaryRptNewMapper.getCustomerOrderQtyMonthly(systemId, customerId, yearmonth);
        RPTCustomerChargeSummaryMonthlyEntity financeMonthly = customerChargeSummaryRptNewMapper.getCustomerFinanceMonthly(systemId, customerId, yearmonth);
        RPTCustomerChargeSummaryMonthlyEntity result = new RPTCustomerChargeSummaryMonthlyEntity();
        result.setCustomerId(customerId);
        result.setYearmonth(yearmonth);
        if (orderQtyMonthly != null) {
            result.setLastMonthUncompletedQty(orderQtyMonthly.getLastMonthUncompletedQty());
            result.setNewQty(orderQtyMonthly.getNewQty());
            result.setCompletedQty(orderQtyMonthly.getCompletedQty());
            result.setReturnedQty(orderQtyMonthly.getReturnedQty());
            result.setCancelledQty(orderQtyMonthly.getCancelledQty());
            result.setUncompletedQty(orderQtyMonthly.getUncompletedQty());
        }
        if (financeMonthly != null) {
            result.setLastMonthBalance(financeMonthly.getLastMonthBalance());
            result.setRechargeAmount(financeMonthly.getRechargeAmount());
            result.setCompletedOrderCharge(financeMonthly.getCompletedOrderCharge());
            result.setWriteOffCharge(financeMonthly.getWriteOffCharge());
            result.setTimelinessCharge(financeMonthly.getTimelinessCharge());
            result.setUrgentCharge(financeMonthly.getUrgentCharge());
            result.setPraiseFee(financeMonthly.getPraiseFee());
            result.setBalance(financeMonthly.getBalance());
            result.setBlockAmount(financeMonthly.getBlockAmount());
        }
        return result;
    }




    //endregion 从Web数据库获取客户的工单数量与消费金额数据


    //endregion

    //region 辅助方法

    private int generateYearMonth(Date date) {
        int selectedYear = DateUtils.getYear(date);
        int selectedMonth = DateUtils.getMonth(date);
        return generateYearMonth(selectedYear, selectedMonth);
    }

    private int generateYearMonth(int selectedYear, int selectedMonth) {
        return StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
    }

    //endregion 辅助方法


    /**
     * 导出 客户对账单
     *
     * @return
     */
    public SXSSFWorkbook exportCustomerChargeRptNew(String searchConditionJson, String reportTitle) {
        RPTCustomerChargeSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerChargeSearch.class);
        RPTCustomerChargeSummaryMonthlyEntity item = getCustomerChargeSummaryNew(searchCondition);
        SXSSFWorkbook xBook = null;
        try {
            long customerId = item.getCustomerId();

            RPTCustomer customer = msCustomerService.get(customerId);

            item.setCustomer(customer == null ? new RPTCustomer() : customer);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet("消费汇总");
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_20);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, item.getCustomer().getName() + "对账单");
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 8));


            Row firstRow = xSheet.createRow(rowIndex++);
            firstRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

            ExportExcel.createCell(firstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "提供服务公司：广东快可立服务有限公司");
            xSheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 7));

            CellRangeAddress region = new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), 0, 7);
            ExportExcel.setRegionBorder(region, xSheet, xBook);

            ExportExcel.createCell(firstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "所属账期:" + searchCondition.getSelectedYear() + "年" + searchCondition.getSelectedMonth() + "月");


            Row firsHeaderRow = xSheet.createRow(rowIndex++);
            firsHeaderRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(firsHeaderRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月未完成单");
            ExportExcel.createCell(firsHeaderRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月下单");
            ExportExcel.createCell(firsHeaderRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单");
            ExportExcel.createCell(firsHeaderRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退单");
            ExportExcel.createCell(firsHeaderRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月取消单");
            ExportExcel.createCell(firsHeaderRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月未完成单");
            ExportExcel.createCell(firsHeaderRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            ExportExcel.createCell(firsHeaderRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
            ExportExcel.createCell(firsHeaderRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

            Row firstDataRow = xSheet.createRow(rowIndex++);
            firstDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(firstDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getLastMonthUncompletedQty());
            ExportExcel.createCell(firstDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getNewQty());
            ExportExcel.createCell(firstDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCompletedQty());
            ExportExcel.createCell(firstDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getReturnedQty());
            ExportExcel.createCell(firstDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCancelledQty());
            ExportExcel.createCell(firstDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUncompletedQty());
            ExportExcel.createCell(firstDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(firstDataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(firstDataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");

            Row secondHeaderRow = xSheet.createRow(rowIndex++);
            secondHeaderRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondHeaderRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月消费余额");
            ExportExcel.createCell(secondHeaderRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月充值");
            ExportExcel.createCell(secondHeaderRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(secondHeaderRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "对帐差异单（本期退补款）");
            ExportExcel.createCell(secondHeaderRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月时效费");
            ExportExcel.createCell(secondHeaderRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月加急费");
            ExportExcel.createCell(secondHeaderRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月好评费");
            ExportExcel.createCell(secondHeaderRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月消费余额");
            ExportExcel.createCell(secondHeaderRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完工冻结金额");

            Row secondDataRow = xSheet.createRow(rowIndex++);
            secondDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(secondDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getLastMonthBalance());
            ExportExcel.createCell(secondDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getRechargeAmount());
            ExportExcel.createCell(secondDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCompletedOrderCharge());
            ExportExcel.createCell(secondDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getWriteOffCharge());
            ExportExcel.createCell(secondDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getTimelinessCharge());
            ExportExcel.createCell(secondDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getUrgentCharge());
            ExportExcel.createCell(secondDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getPraiseFee());
            ExportExcel.createCell(secondDataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBalance());
            ExportExcel.createCell(secondDataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBlockAmount());


            StringBuilder stringBuilder = new StringBuilder();

            Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
            String dayString = DateUtils.formatDate(DateUtils.getLastDayOfMonth(queryDate), "yyyy-MM-dd");
            stringBuilder.append("截止到" + dayString + "," + item.getCustomer().getName());
            stringBuilder.append(item.getBalance() > 0 ? "在广东快可立家电服务有限公司余额为" : "欠广东快可立家电服务有限公司服务款");

            double money = Math.abs(item.getBalance());
            String s = String.valueOf(money);
            CurrencyUtil nf = new CurrencyUtil(s);
            String bigMoney = nf.Convert();

            stringBuilder.append(String.format("%.2f", money) + "元 ");
            stringBuilder.append("（大写：");
            stringBuilder.append(bigMoney);
            stringBuilder.append("）");
            // 截止到2015年3月31号易品购商贸公司欠广东快可立家电服务有限公司服务款18445元（大写：壹万捌仟肆佰肆拾伍元正）

            rowIndex = rowIndex + 2;
            Row remarkRow = xSheet.createRow(rowIndex++);

            ExportExcel.createCell(remarkRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, stringBuilder.toString());
            xSheet.addMergedRegion(new CellRangeAddress(remarkRow.getRowNum(), remarkRow.getRowNum(), 0, 7));

            Row last1Row = xSheet.createRow(rowIndex++);
            ExportExcel.createCell(last1Row, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "确认（单位盖章）:");
            Row last2Row = xSheet.createRow(rowIndex++);
            ExportExcel.createCell(last2Row, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "日期:");

            //添加完工单Sheet
            List<RPTCompletedOrderEntity> customerCompletedOrdersRptData = ctCompletedOrderRptService.getCompletedOrderList(searchCondition);
            if (customerCompletedOrdersRptData.size() < 2000) {
                ctCompletedOrderRptService.addCustomerChargeCompleteRptSheet(xBook, xStyle, customerCompletedOrdersRptData);
            } else {
                ctCompletedOrderRptService.addCustomerChargeCompleteRptSheetMore2000(xBook, xStyle, customerCompletedOrdersRptData);
            }
            //添加退单/取消单Sheet
            List<RPTCancelledOrderEntity> returnedOrderDetail = ctCancelledOrderRptService.getCancelledOrder(searchCondition);

            if (returnedOrderDetail.size() < 2000) {
                ctCancelledOrderRptService.addCustomerChargeReturnCancelRptSheet(xBook, xStyle, returnedOrderDetail);
            } else {
                ctCancelledOrderRptService.addCustomerChargeReturnCancelRptSheetMore2000(xBook, xStyle, returnedOrderDetail);
            }

            //添加退补单Sheet
            List<RPTCustomerWriteOffEntity> WriteOffData = ctCustomerWriteOffRptService.getCustomerWriteOffList(searchCondition);
            if (WriteOffData.size() < 2000) {
                ctCustomerWriteOffRptService.addCustomerWriteOffRptSheet(xBook, xStyle, WriteOffData);
            } else {
                ctCustomerWriteOffRptService.addCustomerWriteOffRptSheetMore2000(xBook, xStyle, WriteOffData);
            }

        } catch (Exception e) {
            log.error("【CustomerChargeSummaryRptNewService.exportCustomerChargeRptNew】客户对账单写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;

    }

}
