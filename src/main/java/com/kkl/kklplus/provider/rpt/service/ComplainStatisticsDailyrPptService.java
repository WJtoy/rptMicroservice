package com.kkl.kklplus.provider.rpt.service;


import com.kkl.kklplus.entity.rpt.RPTComplainStatisticsDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.mapper.ComplainStatisticsDailyMapper;
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
import java.text.NumberFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ComplainStatisticsDailyrPptService extends RptBaseService{

    @Resource
    private ComplainStatisticsDailyMapper complainStatisticsDailyMapper;

    public List<RPTComplainStatisticsDailyEntity> getComplainStatisticsDailyList(RPTComplainStatisticsDailySearch search) {
        List<RPTComplainStatisticsDailyEntity> list = new ArrayList<>();
        search.setSystemId(RptCommonUtils.getSystemId());
        Date startDate  = new Date(search.getStartDate());

       List<RPTComplainStatisticsDailyEntity> dayComplainSums = complainStatisticsDailyMapper.getDayComplainSumNew(search);



        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        int dayComplain = 0;
        int dayTheTotalOrder = 0;

        Map<Integer, RPTComplainStatisticsDailyEntity> dayComplainSumMap = dayComplainSums.stream().collect(Collectors.toMap(RPTComplainStatisticsDailyEntity::getDayIndex,  Function.identity(), (key1, key2) -> key2));


        String key;
        String yearStr;
        String monthStr;
        String daysStr;
        for (int i = 1; i <= 31; i++) {
            Integer keyInt = StringUtils.toInteger(DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd"));
            RPTComplainStatisticsDailyEntity complainStatisticsRptEntity = new RPTComplainStatisticsDailyEntity();
            key = String.valueOf(keyInt);
            yearStr = key.substring(0, 4);
            monthStr = key.substring(4, 6);
            daysStr = key.substring(6, 8);

            RPTComplainStatisticsDailyEntity dayTheTotalOrderSum = dayComplainSumMap.get(keyInt);
            if (dayTheTotalOrderSum == null) {
                complainStatisticsRptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + daysStr);
            } else {
                complainStatisticsRptEntity.setOrderCreateDate(yearStr + "-" + monthStr + "-" + daysStr);
                complainStatisticsRptEntity.setDayTheTotalOrder(dayTheTotalOrderSum.getDayTheTotalOrder());
                complainStatisticsRptEntity.setDayComplainSum(dayTheTotalOrderSum.getDayComplainSum());

                int total = complainStatisticsRptEntity.getDayTheTotalOrder();
                countComplainStatisticsRptDayComplainSumRate(numberFormat, complainStatisticsRptEntity, total);

                dayComplain += complainStatisticsRptEntity.getDayComplainSum();
                dayTheTotalOrder += complainStatisticsRptEntity.getDayTheTotalOrder();
            }
            list.add(complainStatisticsRptEntity);

        }

        list = list.stream().sorted(Comparator.comparing(RPTComplainStatisticsDailyEntity::getOrderCreateDate)).collect(Collectors.toList());
        //合计
        RPTComplainStatisticsDailyEntity rptEntity = new RPTComplainStatisticsDailyEntity();
        rptEntity.setDayComplainSum(dayComplain);
        rptEntity.setDayTheTotalOrder(dayTheTotalOrder);

        Integer total = rptEntity.getDayTheTotalOrder();
        countComplainStatisticsRptDayComplainSumRate(numberFormat, rptEntity, total);
        list.add(rptEntity);

        return list;

    }

    /**
     * 计算投诉报表每日投诉比率
     *
     * @param numberFormat
     * @param complainStatisticsRptEntity
     * @param total
     */
    private void countComplainStatisticsRptDayComplainSumRate(NumberFormat numberFormat, RPTComplainStatisticsDailyEntity complainStatisticsRptEntity, Integer total) {
        if (total != null && total != 0) {
            int complainOrder = complainStatisticsRptEntity.getDayComplainSum();
            if (complainOrder > 0) {
                float complainOrderRateInt = (float) complainOrder / total * 100;
                String complainOrderRate = numberFormat.format(complainOrderRateInt);
                complainStatisticsRptEntity.setDayComplainSumRate(complainOrderRate);
                complainStatisticsRptEntity.setDayComplainSumRateInt(complainOrderRateInt);
            }
        }
    }

    public Map<String, Object> getComplainStatisticsDailyChartList(RPTComplainStatisticsDailySearch search) {
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        Map<String, Object> map = new HashMap<>();
        List<RPTComplainStatisticsDailyEntity> entityList = getComplainStatisticsDailyList(search);

        if (entityList == null || entityList.size() <= 0) {
            return map;
        }

        List<String> createDates = new ArrayList<>();
        List<String> orderCreateDates = new ArrayList<>();

        List<String> strDayComplainSum = new ArrayList<>();
        List<String> strDayComplainSumRate = new ArrayList<>();
        List<Float> intDayComplainSumRate = new ArrayList<>();

        for (RPTComplainStatisticsDailyEntity entity : entityList) {
            if (entity.getOrderCreateDate() != null) {
                orderCreateDates.add(entity.getOrderCreateDate().substring(5));
                createDates.add(entity.getOrderCreateDate().substring(5));
                strDayComplainSum.add(entity.getDayComplainSum().toString());
                strDayComplainSumRate.add(entity.getDayComplainSumRate());
                intDayComplainSumRate.add(entity.getDayComplainSumRateInt());
            }
        }
        int rate = (int) Math.ceil(Collections.max(intDayComplainSumRate));
        if (rate > 5) {
            rate = (rate + 4) / 5 * 5;
        } else {
            rate = 5;
        }
        map.put("createDates", createDates);
        map.put("orderCreateDates", orderCreateDates);
        map.put("rate", rate);
        map.put("strDayComplainSum", strDayComplainSum);
        map.put("strDayComplainSumRate", strDayComplainSumRate);


        return map;

    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
         searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getEndDate() != null && searchCondition.getStartDate() != null) {
            Integer rowCount = complainStatisticsDailyMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 每日投诉统计导出
     *
     * @return
     */

    public SXSSFWorkbook ComplainStatisticsDailyExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTComplainStatisticsDailySearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTComplainStatisticsDailySearch.class);
            List<RPTComplainStatisticsDailyEntity> list = getComplainStatisticsDailyList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;
            //===============绘制标题======================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单数量");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉数量");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉比率");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTComplainStatisticsDailyEntity entity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    if (dataRowIndex == list.size() - 1) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                    } else {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getOrderCreateDate());
                    }
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getDayTheTotalOrder());
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getDayComplainSum());
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, entity.getDayComplainSumRate() + "%");

                }
            }

        } catch (Exception e) {
            log.error("【ComplainStatisticsDailyrPptService.ComplainStatisticsDailyExport】每日投诉统计报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
