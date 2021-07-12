package com.kkl.kklplus.provider.rpt.service;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTSMSQtyStatisticsEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTSMSQtyStatisticsSearch;
import com.kkl.kklplus.entity.sys.SysSMSTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.SmsQtyEntity;
import com.kkl.kklplus.provider.rpt.mapper.SMSQtyStatisticsMapper;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSSmsQtyFeign;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateFormatUtils;
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

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class SMSQtyStatisticsService extends RptBaseService{

    @Autowired
    private MSSmsQtyFeign msSmsQtyFeign;

    @Resource
    private SMSQtyStatisticsMapper smsQtyStatisticsMapper;

    /**
     * 根据时间查询
     *
     * @return
     */
    public List<RPTSMSQtyStatisticsEntity> getMessageNumberRpt(RPTSMSQtyStatisticsSearch search) {
        Date queryDate = DateUtils.getDate(search.getSelectedYear(), search.getSelectedMonth(), 1);
        Date startDate = DateUtils.getStartDayOfMonth(queryDate);
        Date endDate = DateUtils.getLastDayOfMonth(queryDate);
        Long startDt = startDate.getTime();
        Long endDt = endDate.getTime();
        int systemId = RptCommonUtils.getSystemId();
        List<RPTSMSQtyStatisticsEntity> lists = new ArrayList<>();
        int dayNum = 0;
        int createdNum = 0;
        int plannedNum = 0;
        int acceptedAppNum = 0;
        int servicePointNum = 0;
        int pendingNum = 0;
        int pendingAppNum = 0;
        int verificationCodeNum = 0;
        int orderDetailPageNum = 0;
        int callBackNum = 0;
        int cancelledNum = 0;
        int dayIndex;
        String stringDate;
        String strDayIndex;
        String yearStr;
        String monthStr;
        String dayStr;
        List<RPTSMSQtyStatisticsEntity> list = smsQtyStatisticsMapper.getMessageType(systemId,startDt, endDt);

        for (RPTSMSQtyStatisticsEntity smsQtyRptEntity : list) {
            strDayIndex = smsQtyRptEntity.getDayIndex().toString();
            yearStr = strDayIndex.substring(0, 4);
            monthStr = strDayIndex.substring(4, 6);
            dayStr = strDayIndex.substring(6, 8);
            smsQtyRptEntity.setSendDate(yearStr + "-" + monthStr + "-" + dayStr);
            createdNum += smsQtyRptEntity.getCreated();
            plannedNum += smsQtyRptEntity.getPlanned();
            acceptedAppNum += smsQtyRptEntity.getAcceptedApp();
            servicePointNum += smsQtyRptEntity.getServicePoint();
            pendingNum += smsQtyRptEntity.getPending();
            pendingAppNum += smsQtyRptEntity.getPendingApp();
            verificationCodeNum += smsQtyRptEntity.getVerificationCode();
            orderDetailPageNum += smsQtyRptEntity.getOrderDetailPage();
            callBackNum += smsQtyRptEntity.getCallBack();
            cancelledNum += smsQtyRptEntity.getCancelled();
            dayNum += smsQtyRptEntity.getDayNum();
            lists.add(smsQtyRptEntity);
        }
        //从缓存中取出当天数据
        Date date = new Date();
        int year = DateUtils.getYear(queryDate);
        int newYear = DateUtils.getYear(date);
        int month = DateUtils.getMonth(queryDate);
        int newMonth = DateUtils.getMonth(date);

        RPTSMSQtyStatisticsEntity entity = new RPTSMSQtyStatisticsEntity();
        if (year == newYear && month == newMonth) {
            strDayIndex = DateFormatUtils.format(date, "yyyyMMdd");
            stringDate = DateFormatUtils.format(date, "yyyy-MM-dd");
            MSResponse<Map<Integer, Long>> map = msSmsQtyFeign.shortMessageCache(stringDate);
            if (map.getCode() == 0) {
                Map<Integer, Long> longMap = map.getData();
                if (longMap.size() != 0) {
                    for (Integer key : longMap.keySet()) {
                        if (key == SysSMSTypeEnum.ORDER_CREATED.getValue()) {
                            entity.setCreated(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.ORDER_PLANNED.getValue()) {
                            entity.setPlanned(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.ORDER_ACCEPTED_APP.getValue()) {
                            entity.setAcceptedApp(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.ORDER_PLANNED_SERVICE_POINT.getValue()) {
                            entity.setServicePoint(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.ORDER_PENDING.getValue()) {
                            entity.setPending(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.ORDER_PENDING_APP.getValue()) {
                            entity.setPendingApp(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.VERIFICATION_CODE.getValue()) {
                            entity.setVerificationCode(longMap.get(key).hashCode());
                        } else if (key == SysSMSTypeEnum.KEFU_ORDERDETAIL_PAGE_TRIGGER.getValue()) {
                            entity.setOrderDetailPage(longMap.get(key).intValue());
                        } else if (key == SysSMSTypeEnum.CALL_BACK.getValue()) {
                            entity.setCallBack(longMap.get(key).intValue());
                        } else if(key == SysSMSTypeEnum.ORDER_CANCELLED.getValue()){
                            entity.setCancelled(longMap.get(key).intValue());
                        }
                    }
                    if (entity.getDayNum() != 0) {
                        dayIndex = Integer.parseInt(strDayIndex);
                        entity.setDayIndex(dayIndex);
                        entity.setSendDate(stringDate);
                        lists.add(entity);
                    }
                }
            } else {
                log.error("[SMSQtyStatisticsService.getMessageNumberRpt]调用短信微服务接口失败");
            }
        }
        if(lists.size() <= 0){
            return lists;
        }
        //合计
        RPTSMSQtyStatisticsEntity smsQtyRptEntity = new RPTSMSQtyStatisticsEntity();
        smsQtyRptEntity.setCreated(createdNum + entity.getCreated());
        smsQtyRptEntity.setPlanned(plannedNum + entity.getPlanned());
        smsQtyRptEntity.setAcceptedApp(acceptedAppNum + entity.getAcceptedApp());
        smsQtyRptEntity.setServicePoint(servicePointNum + entity.getServicePoint());
        smsQtyRptEntity.setPending(pendingNum + entity.getPending());
        smsQtyRptEntity.setPendingApp(pendingAppNum + entity.getPendingApp());
        smsQtyRptEntity.setVerificationCode(verificationCodeNum + entity.getVerificationCode());
        smsQtyRptEntity.setOrderDetailPage(orderDetailPageNum + entity.getOrderDetailPage());
        smsQtyRptEntity.setCallBack(callBackNum + entity.getCallBack());
        smsQtyRptEntity.setCancelled(cancelledNum + entity.getCancelled());
        smsQtyRptEntity.setDayNum(dayNum + entity.getDayNum());
        lists.add(smsQtyRptEntity);
        return lists;
    }

    public Map<String, Object> turnToChartInformation(RPTSMSQtyStatisticsSearch search) {
        Map<String, Object> map = new HashMap<>();
        List<RPTSMSQtyStatisticsEntity> entityList = getMessageNumberRpt(search);
        if (entityList == null || entityList.size() <= 0) {
            return map;
        }
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);

        List<String> sendDateList = new ArrayList<>();
        List<String> plannedList = new ArrayList<>();
        List<String> acceptedAppList = new ArrayList<>();
        List<String> pendingList = new ArrayList<>();
        List<String> pendingApps = new ArrayList<>();
        List<String> verificationCodes = new ArrayList<>();
        List<String> orderDetailPages = new ArrayList<>();
        List<String> callBacks = new ArrayList<>();
        List<String> cancelleds = new ArrayList<>();

        List<String> plannedRates = new ArrayList<>();
        List<String> acceptedAppRates = new ArrayList<>();
        List<String> pendingRates = new ArrayList<>();
        List<String> pendingAppRates = new ArrayList<>();
        List<String> verificationCodeRates = new ArrayList<>();
        List<String> orderDetailPageRates = new ArrayList<>();
        List<String> callBackRates = new ArrayList<>();
        List<String> cancelledRates = new ArrayList<>();


        Map<String, String> plannedMap = new HashMap<>();
        Map<String, String> acceptedAppMap = new HashMap<>();
        Map<String, String> pendingMap = new HashMap<>();
        Map<String, String> pendingAppMap = new HashMap<>();
        Map<String, String> verificationCodeMap = new HashMap<>();
        Map<String, String> orderDetailPageMap = new HashMap<>();
        Map<String, String> callBackMap = new HashMap<>();
        Map<String, String> cancelledMap = new HashMap<>();

        RPTSMSQtyStatisticsEntity smsQtyRptEntity = entityList.get(entityList.size() - 1);
        plannedMap.put("value", smsQtyRptEntity.getPlanned().toString());
        plannedMap.put("name", "派单");
        acceptedAppMap.put("value", smsQtyRptEntity.getAcceptedApp().toString());
        acceptedAppMap.put("name", "APP接单");
        pendingMap.put("value", smsQtyRptEntity.getPending().toString());
        pendingMap.put("name", "客服预约");
        pendingAppMap.put("value", smsQtyRptEntity.getPendingApp().toString());
        pendingAppMap.put("name", "网点预约");
        verificationCodeMap.put("value", smsQtyRptEntity.getVerificationCode().toString());
        verificationCodeMap.put("name", "验证码");
        orderDetailPageMap.put("value", smsQtyRptEntity.getOrderDetailPage().toString());
        orderDetailPageMap.put("name", "客服工单详情界面");
        callBackMap.put("value", smsQtyRptEntity.getCallBack().toString());
        callBackMap.put("name", "短信回访");
        cancelledMap.put("value", smsQtyRptEntity.getCancelled().toString());
        cancelledMap.put("name", "订单取消");

        List<Map<String, String>> mapList = new ArrayList<>();

        for (RPTSMSQtyStatisticsEntity entity : entityList) {
            if (entity.getDayIndex() != null) {
                int dayNum = entity.getDayNum();
                String sendDate = entity.getSendDate();
                sendDateList.add(sendDate.substring(8));
                plannedList.add(entity.getPlanned().toString());
                acceptedAppList.add(entity.getAcceptedApp().toString());
                pendingList.add(entity.getPending().toString());
                pendingApps.add(entity.getPendingApp().toString());
                verificationCodes.add(entity.getVerificationCode().toString());
                orderDetailPages.add(entity.getOrderDetailPage().toString());
                callBacks.add(entity.getCallBack().toString());
                cancelleds.add(entity.getCancelled().toString());

                if (dayNum != 0) {
                    int planned = entity.getPlanned();
                    int acceptedApp = entity.getAcceptedApp();
                    int pending = entity.getPending();
                    int pendingApp = entity.getPendingApp();
                    int verificationCode = entity.getVerificationCode();
                    int orderDetailPage = entity.getOrderDetailPage();
                    int callBack = entity.getCallBack();
                    int cancelled = entity.getCancelled();
                    if (planned != 0) {
                        String plannedRate = numberFormat.format((float) planned / dayNum * 100);
                        plannedRates.add(plannedRate);
                    }else {
                        plannedRates.add("0");
                    }
                    if (acceptedApp != 0) {
                        String acceptedAppRate = numberFormat.format((float) acceptedApp / dayNum * 100);
                        acceptedAppRates.add(acceptedAppRate);
                    }else {
                        acceptedAppRates.add("0");
                    }
                    if (pending != 0) {
                        String pendingRate = numberFormat.format((float) pending / dayNum * 100);
                        pendingRates.add(pendingRate);
                    }else {
                        pendingRates.add("0");
                    }
                    if(pendingApp != 0){
                        String pendingAppRate = numberFormat.format((float) pendingApp / dayNum * 100);
                        pendingAppRates.add(pendingAppRate);
                    }else {
                        pendingAppRates.add("0");
                    }
                    if(verificationCode != 0){
                        String verificationCodeRate = numberFormat.format((float) verificationCode / dayNum * 100);
                        verificationCodeRates.add(verificationCodeRate);
                    }else {
                        verificationCodeRates.add("0");
                    }
                    if(orderDetailPage != 0){
                        String orderDetailPageRate = numberFormat.format((float) orderDetailPage / dayNum * 100);
                        orderDetailPageRates.add(orderDetailPageRate);
                    }else {
                        orderDetailPageRates.add("0");
                    }
                    if(callBack != 0){
                        String callBackRate = numberFormat.format((float) callBack / dayNum * 100);
                        callBackRates.add(callBackRate);
                    }else {
                        callBackRates.add("0");
                    }
                    if(cancelled != 0){
                        String cancelledRate = numberFormat.format((float) cancelled / dayNum * 100);
                        cancelledRates.add(cancelledRate);
                    }else {
                        cancelledRates.add("0");
                    }
                }
            }
        }

        mapList.add(plannedMap);
        mapList.add(acceptedAppMap);
        mapList.add(pendingMap);
        mapList.add(pendingAppMap);
        mapList.add(verificationCodeMap);
        mapList.add(orderDetailPageMap);
        mapList.add(callBackMap);
        mapList.add(cancelledMap);
        map.put("sendDateList",sendDateList);
        map.put("plannedList",plannedList);
        map.put("acceptedAppList",acceptedAppList);
        map.put("pendingList",pendingList);
        map.put("pendingApps",pendingApps);
        map.put("verificationCodes",verificationCodes);
        map.put("orderDetailPages",orderDetailPages);
        map.put("callBacks",callBacks);
        map.put("cancelleds",cancelleds);

        map.put("plannedRates",plannedRates);
        map.put("acceptedAppRates",acceptedAppRates);
        map.put("pendingRates",pendingRates);
        map.put("pendingAppRates",pendingAppRates);
        map.put("verificationCodeRates",verificationCodeRates);
        map.put("orderDetailPageRates",orderDetailPageRates);
        map.put("callBackRates",callBackRates);
        map.put("cancelledRates",cancelledRates);
        map.put("mapList",mapList);

        return map;
    }


    /**
     * 填充短信类型数据
     * @param messageNumber
     * @param date
     * @param smsQty
     */
    private void paddingDataSmsQty(SmsQtyEntity messageNumber, Date date, Map<Integer, Long> smsQty) {
        int dayIndex;
        int systemId = RptCommonUtils.getSystemId();
        String strDayIndex;
        for (Map.Entry<Integer, Long> entry : smsQty.entrySet()) {
            strDayIndex = DateUtils.formatDate(date, "yyyyMMdd");
            dayIndex = Integer.parseInt(strDayIndex);
            messageNumber.setSystemId(systemId);
            messageNumber.setSmsType(entry.getKey());
            messageNumber.setSmsQty(entry.getValue().intValue());
            messageNumber.setSendDate(date.getTime());
            messageNumber.setDayIndex(dayIndex);
            smsQtyStatisticsMapper.insert(messageNumber);
        }
    }

    /**
     * 检查当前时间前两天的缓存与数据库数量
     *
     * @param
     */
    public void checkEveShortMessageQty() {
        SmsQtyEntity messageNumber = new SmsQtyEntity();
        Date date = DateUtils.addDays(new Date(), -2);
        Date startDate = DateUtils.getDateStart(date);//前一天的0：00
        Long endDt = DateUtils.getDateEnd(date).getTime();//前一天的23：59
        Long startDt = startDate.getTime();
        int systemId = RptCommonUtils.getSystemId();
        try {
            List<SmsQtyEntity> list = smsQtyStatisticsMapper.getSmsQtyRpt(systemId,startDt, endDt);
            String stringDate = DateFormatUtils.format(date, "yyyy-MM-dd");
            MSResponse<Map<Integer, Long>> map = msSmsQtyFeign.shortMessageCache(stringDate);
            if (map.getCode() == 0) {
                Map<Integer, Long> smsQty = map.getData();
                for (SmsQtyEntity entity : list) {
                    if (smsQty.get(entity.getSmsType()) == null) {
                        log.error("[SmsQtyRptService.inspect]检查当前时间前2天的数据缓存数据有空异常");
                    } else if (entity.getSmsQty().longValue() != smsQty.get(entity.getSmsType())) {
                        entity.setSystemId(systemId);
                        entity.setSmsQty(smsQty.get(entity.getSmsType()).intValue());
                        smsQtyStatisticsMapper.update(entity);
                    }
                    smsQty.remove(entity.getSmsType());
                }
                if (smsQty.size() > 0) {
                    paddingDataSmsQty(messageNumber,startDate, smsQty);
                }
            }
        } catch (Exception e) {
            log.error("[SmsQtyRptService.inspect]检查当前时间前2天的数据", e);
        }

    }

    /**
     * 写入当前一天数据
     */
    public void writeYesterdayMessage(Date date) {
        SmsQtyEntity messageNumber = new SmsQtyEntity();
        String strDayIndex = DateFormatUtils.format(date, "yyyy-MM-dd");
        MSResponse<Map<Integer, Long>> map = msSmsQtyFeign.shortMessageCache(strDayIndex);
        if (map.getCode() == 0) {
            Map<Integer, Long> smsQty = map.getData();
            if (smsQty.size() != 0) {
                paddingDataSmsQty(messageNumber,date, smsQty);
            }
        }
    }
    public void saveSMSQtyStatisticsToRptDB(Date date){
        int systemId = RptCommonUtils.getSystemId();
        Date endDate = DateUtils.getEndOfDay(date);
        Date startDate = DateUtils.getStartOfDay(date);
        List<SmsQtyEntity> sysSmsQtyList = smsQtyStatisticsMapper.getSysSmsQtyData(startDate, endDate);
        String strDayIndex;
        for(SmsQtyEntity entity : sysSmsQtyList){
            entity.setSystemId(systemId);
            entity.setSendDate(entity.getSmsSendDate().getTime());
            strDayIndex = DateFormatUtils.format(entity.getSmsSendDate(), "yyyyMMdd");
            entity.setDayIndex(Integer.parseInt(strDayIndex));
            smsQtyStatisticsMapper.insert(entity);
        }
    }
    public void deleteSMSQtyStatisticsRptDB(Date date){
        if (date != null) {
            Date beginDate = DateUtils.getDateStart(date);
            Date endDate = DateUtils.getDateEnd(date);
            int systemId = RptCommonUtils.getSystemId();
            smsQtyStatisticsMapper.delete(systemId, beginDate.getTime(), endDate.getTime());
        }
    }
    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {
                Date beginDate = new Date(beginDt);
                Date endDate = DateUtils.getEndOfDay(new Date(endDt));
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            saveSMSQtyStatisticsToRptDB(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            break;
                        case UPDATE:
                            deleteSMSQtyStatisticsRptDB(beginDate);
                            saveSMSQtyStatisticsToRptDB(beginDate);
                            break;
                        case DELETE:
                            deleteSMSQtyStatisticsRptDB(beginDate);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("KeFuCompleteTimeRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTSMSQtyStatisticsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSMSQtyStatisticsSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getSelectedYear() != null && searchCondition.getSelectedMonth() != null) {
            Date queryDate = DateUtils.getDate(searchCondition.getSelectedYear(), searchCondition.getSelectedMonth(), 1);
            Date startDate = DateUtils.getStartDayOfMonth(queryDate);
            Date endDate = DateUtils.getLastDayOfMonth(queryDate);
            int systemId = RptCommonUtils.getSystemId();
            Long startDt = startDate.getTime();
            Long endDt = endDate.getTime();
            Integer rowCount = smsQtyStatisticsMapper.hasReportData(systemId,startDt,endDt);
            result = rowCount > 0;
        }
        return result;
    }
    /**
     * 报表导出
     */
    public SXSSFWorkbook messageNumRptExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTSMSQtyStatisticsSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTSMSQtyStatisticsSearch.class);
            List<RPTSMSQtyStatisticsEntity> list = getMessageNumberRpt(searchCondition);
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
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
            //=================绘制表头行=========================
            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "日期");
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单");
            ExportExcel.createCell(headFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "APP接单");
            ExportExcel.createCell(headFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点转派");
            ExportExcel.createCell(headFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服设置停滞");
            ExportExcel.createCell(headFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点设置停滞");
            ExportExcel.createCell(headFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "验证码");
            ExportExcel.createCell(headFirstRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服工单详情界面");
            ExportExcel.createCell(headFirstRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "短信回访");
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "订单取消");
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "总短信数量");
            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            if (list != null && list.size() > 0) {
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTSMSQtyStatisticsEntity smsQtyRptEntity = list.get(dataRowIndex);
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    if (dataRowIndex == list.size() - 1) {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "合计");
                    } else {
                        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getSendDate());
                    }
                    if (smsQtyRptEntity.getDayNum() != 0) {
                        ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getPlanned());
                        ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getAcceptedApp());
                        ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getServicePoint());
                        ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getPending());
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getPendingApp());
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getVerificationCode());
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getOrderDetailPage());
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getCallBack());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getCancelled());
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, smsQtyRptEntity.getDayNum());
                    }

                }
            }

        } catch (Exception e) {
            log.error("【SMSQtyStatisticsService.messageNumRptExport】短信数量统计报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
