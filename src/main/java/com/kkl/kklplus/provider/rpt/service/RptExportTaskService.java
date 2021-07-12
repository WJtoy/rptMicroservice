package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.kkl.kklplus.entity.rpt.RPTExportTaskEntity;
import com.kkl.kklplus.entity.rpt.RPTExportTaskSearch;
import com.kkl.kklplus.entity.rpt.common.RPTErrorCode;
import com.kkl.kklplus.entity.rpt.common.RPTReportEnum;
import com.kkl.kklplus.entity.rpt.common.RPTReportExportStatusEnum;
import com.kkl.kklplus.entity.rpt.common.RPTReportTypeEnum;
import com.kkl.kklplus.entity.rpt.exception.RPTBaseException;
import com.kkl.kklplus.provider.rpt.customer.service.*;
import com.kkl.kklplus.provider.rpt.mapper.RptExportTaskMapper;
import com.kkl.kklplus.provider.rpt.mq.receiver.RPTExportReportTaskMQReceiver;
import com.kkl.kklplus.provider.rpt.mq.sender.RPTExportReportTaskMQSender;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.RptExcelFileUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import com.kkl.kklplus.starter.redis.utils.RedisDefaultDbNewUtils;
import lombok.extern.slf4j.Slf4j;
import om.kkl.kklplus.entity.rpt.mq.pb.MQRPTExportTaskMessage;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.javatuples.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;

/**
 * 报表导出服务
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class RptExportTaskService {

    @Resource
    RptExportTaskMapper rptExportTaskMapper;

    @Autowired
    RPTExportReportTaskMQSender rptExportReportTaskMQSender;
    @Autowired
    RPTExportReportTaskMQReceiver rptExportReportTaskMQReceiver;

    @Autowired
    private CustomerOrderDailyRptService customerOrderDailyRptService;

    @Autowired
    private SpecialChargeAreaRptService specialChargeAreaRptService;

    @Autowired
    private TravelChargeRankRptService travelChargeRankRptService;

    @Autowired
    private CustomerChargeSummaryRptService customerChargeSummaryRptService;

    @Autowired
    private UncompletedOrderRptService uncompletedOrderRptService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    @Autowired
    private CustomerReminderRptService customerReminderRptService;

    @Autowired
    private ReminderResponseTimeRptService reminderResponseTimeRptService;

    @Autowired
    private CompletedOrderDetailsService completedOrderDetailsService;

    @Autowired
    private AreaOrderPlanDailyRptService areaOrderPlanDailyRptService;

    @Autowired
    private KeFuOrderPlanDailyRptService keFuOrderPlanDailyRptService;

    @Autowired
    private CustomerOrderPlanDailyRptService customerOrderPlanDailyRptService;

    @Autowired
    private KeFuOrderCancelledDailyRptService keFuOrderCancelledDailyRptService;

    @Autowired
    private RptServicePonintWriteService rptServicePonintWriteService;

    @Autowired
    private RPTOrderDailyWorkService rptOrderDailyWorkService;

    @Autowired
    private ExploitDetailRptService exploitDetailRptService;

    @Autowired
    private KeFuCompleteTimeRptService keFuCompleteTimeRptService;

    @Autowired
    private CustomerOrderTimeRptService customerOrderTimeRptService;

    @Autowired
    private CustomerMonthPlanDailyRptService customerMonthPlanDailyRptService;

    @Autowired
    private DispatchListInformationRptService dispatchListInformationRptService;

    @Autowired
    private KeFuCompletedMonthRptService keFuCompletedMonthRptService;

    @Autowired
    private SMSQtyStatisticsService smsQtyStatisticsService;

    @Autowired
    private ComplainStatisticsDailyrPptService complainStatisticsDailyrPptService;

    @Autowired
    private CustomerRevenueRptService customerRevenueRptService;

    @Autowired
    private ServicePointInvoiceRptService servicePointInvoiceRptService;

    @Autowired
    private CustomerFinanceRptService customerFinanceRptService;

    @Autowired
    private  ServicePointBalanceRptService servicePointBalanceRptService;

    @Autowired
    private GradeQtyDailyRptService gradeQtyDailyRptService;

    @Autowired
    private CustomerComplainRptService customerComplainRptService;

    @Autowired
    private CustomerPerformanceRptService customerPerformanceRptService;

    @Autowired
    private CustomerReceivableSummaryRptService customerReceivableSummaryRptService;

    @Autowired
    private CustomerChargeSummaryRptNewService customerChargeSummaryRptNewService;

    @Autowired
    private ServicePointChargeRptService servicePointChargeRptService;

    @Autowired
    private RechargeRecordRptService rechargeRecordRptService;

    @Autowired
    private ChargeDailyRptService chargeDailyRptService;

    @Autowired
    private ServicePointCoverageRptService servicePointCoverageRptService;

    @Autowired
    private ServicePointPaymentSummaryRptService servicePointPaymentSummaryRptService;

    @Autowired
    private CustomerRechargeSummaryRptService customerRechargeSummaryRptService;

    @Autowired
    private ComplainRatioDailyRptService complainRatioDailyRptService;

    @Autowired
    private ServicePointBaseInfoRptService servicePointBaseInfoRptService;

    @Autowired
    private AbnormalFinancialReviewRptService abnormalFinancialReviewRptService;

    @Autowired
    private KeFuPraiseDetailsRptService keFuPraiseDetailsRptService;

    @Autowired
    private CustomerPraiseDetailsRptService customerPraiseDetailsRptService;

    @Autowired
    private ServicePointPraiseDetailsRptService servicePointPraiseDetailsRptService;

    @Autowired
    private CustomerFrozenDailyRptService customerFrozenDailyRptService;

    @Autowired
    private FinancialReviewDetailsRptService financialReviewDetailsRptService;

    @Autowired
    private UncompletedQtyRptService uncompletedQtyRptService;

    @Autowired
    private KaReminderResponseTimeRptService kaReminderResponseTimeRptService;

    @Autowired
    private CtCustomerChargeSummaryRptNewService ctCustomerChargeSummaryRptNewService;

    @Autowired
    private CtCustomerOrderDailyRptService ctCustomerOrderDailyRptService;

    @Autowired
    private CtUncompletedOrderRptService ctUncompletedOrderRptService;

    @Autowired
    private CrushCoverageRptService crushCoverageRptService;

    @Autowired
    private TravelCoverageRptService travelCoverageRptService;

    @Autowired
    private RptServicePonintWriteNewService rptServicePonintWriteNewService;

    @Autowired
    private CompletedOrderNewDetailsService completedOrderNewDetailsService;

    @Autowired
    private KeFuCompleteTimeNewRptService keFuCompleteTimeNewRptService;

    @Autowired
    private CustomerOrderMonthRptService customerOrderMonthRptService;

    @Autowired
    private CustomerFrozenMonthRptService customerFrozenMonthRptService;

    @Autowired
    private KeFuAreaRptService keFuAreaRptService;

    @Autowired
    private MasterApplyRptService masterApplyRptService;

    @Autowired
    private CustomerSpecialChargeAreaRptService customerSpecialChargeAreaRptService;

    @Autowired
    private CrushAreaRptService crushAreaRptService;

    @Autowired
    private CancelledOrderNewRptService cancelledOrderNewRptService;

    @Autowired
    private KeFuAverageOrderFeeRptService keFuAverageOrderFeeRptService;

    @Autowired
    private CtCustomerFinanceRptService ctCustomerFinanceRptService;

    @Autowired
    private CtCustomerOrderPlanDailyRptService ctCustomerOrderPlanDailyRptService;

    @Autowired
    private CtCustomerPraiseDetailsRptService ctCustomerPraiseDetailsRptService;

    @Autowired
    private CustomerContractRptService customerContractRptService;

    @Autowired
    private CtCustomerOrderMonthRptService ctCustomerOrderMonthRptService;

    @Autowired
    private CtCustomerFrozenMonthRptService ctCustomerFrozenMonthRptService;

    @Autowired
    private DepositRechargeSummaryRptService depositRechargeSummaryRptService;

    @Autowired
    private EveryDayCompleteService everyDayCompleteService;

    public RPTExportTaskEntity get(Long id) {
        RPTExportTaskEntity taskEntity = null;
        if (id != null) {
            taskEntity = rptExportTaskMapper.get(id);
        }
        return taskEntity;
    }

    @Transactional()
    public Pair<Boolean, String> downloadReportExcel(RPTExportTaskEntity params) {
        Pair<Boolean, String> result = new Pair<>(false, "");
        if (params != null && params.getId() != null && params.getId() > 0
                && params.getLastDownloadBy() != null && params.getLastDownloadBy() > 0) {
            RPTExportTaskEntity task = get(params.getId());
            if (task != null && task.getStatus() >= RPTReportExportStatusEnum.SUCCESS.value
                    && StringUtils.isNotBlank(task.getReportFilePath()) && params.getLastDownloadBy().equals(task.getTaskCreateBy())) {
                String downloadUrl = RptExcelFileUtils.getRptExcelFileHostDir() + task.getReportFilePath();
                task.setStatus(RPTReportExportStatusEnum.DOWNLOADED.value);
                task.setLastDownloadBy(params.getLastDownloadBy());
                task.setLastDownloadByName(StringUtils.toString(params.getLastDownloadByName()));
                task.setLastDownloadDate(System.currentTimeMillis());
                rptExportTaskMapper.updateDownloadInfo(task);
                result = new Pair<>(true, downloadUrl);
            }
        }
        return result;
    }

    /**
     * 创建报表导出任务
     * 参数：reportId、taskCreateBy、searchConditionJson
     */
    @Transactional()
    public void checkRptExportTask(RPTExportTaskEntity params) {
        if (params != null && RPTReportEnum.isReportId(params.getReportId())
                && params.getTaskCreateBy() != null && params.getTaskCreateBy() > 0) {
            String searchConditionJson = StringUtils.trim(StringUtils.toString(params.getSearchConditionJson()));
            if (!hasReportData(params.getReportId(), searchConditionJson)) {
                throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "导出失败：报表没有数据，无需导出");
            }
            Long taskId = getRptExportTaskIdFromRedis(params.getReportId(), params.getTaskCreateBy(), searchConditionJson.hashCode());
            if (taskId != null) {
                throw new RPTBaseException(RPTErrorCode.RPT_EXPORT_TASK_IS_EXISTS, "导出失败：一分钟内不允许重复导出同样的报表，请前往'报表导出->报表下载'功能直接下载");
            }
        } else {
            throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "导出失败：报表参数不全");
        }
    }

    /**
     * 创建报表导出任务
     * 参数：reportId、reportType、taskCreateBy、taskCreateByName、reportTitle、searchConditionJson
     */
    @Transactional()
    public void createRptExportTask(RPTExportTaskEntity params) {
        if (params != null
                && RPTReportEnum.isReportId(params.getReportId())
                && RPTReportTypeEnum.isReportType(params.getReportType())
                && params.getTaskCreateBy() != null && params.getTaskCreateBy() > 0
                && StringUtils.isNotBlank(params.getReportTitle())) {
            params.setTaskCreateByName(StringUtils.toString(params.getTaskCreateByName()));
            params.setTaskCreateDate(System.currentTimeMillis());
            params.setSearchConditionJson(StringUtils.trim(StringUtils.toString(params.getSearchConditionJson())));
            params.setSearchConditionHashcode(params.getSearchConditionJson().hashCode());
//            List<RPTExportTaskEntity> list = rptExportTaskMapper.getTaskListByHashCode(params.getReportId(), params.getTaskCreateDate() - 600 * 1000, params.getTaskCreateBy(), params.getSearchConditionHashcode());
//            if (!list.isEmpty()) {
//                throw new RPTBaseException(RPTErrorCode.RPT_EXPORT_TASK_IS_EXISTS);
//            }
//            if (!hasReportData(params.getReportId(), params.getSearchConditionJson())) {
//                throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "创建报表导出任务失败：报表没有数据，无需导出");
//            }
            Long taskId = getRptExportTaskIdFromRedis(params.getReportId(), params.getTaskCreateBy(), params.getSearchConditionHashcode());
            if (taskId != null) {
                throw new RPTBaseException(RPTErrorCode.RPT_EXPORT_TASK_IS_EXISTS, "一分钟内不允许重复导出同样的报表，请前往'报表中心->报表下载'功能直接下载");
            }
            String quarter = QuarterUtils.getSeasonQuarter(params.getTaskCreateDate());
            params.setQuarter(quarter);
            params.setStatus(RPTReportExportStatusEnum.EXPORTING.value);
            params.setSystemId(RptCommonUtils.getSystemId());
            rptExportTaskMapper.insert(params);
            boolean result = sendRptExportTaskMessage(params);
            if (result) {
                saveExportTaskIdToRedis(params.getId(), params.getReportId(), params.getTaskCreateBy(), params.getSearchConditionHashcode());
            } else {
                throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "创建报表导出任务失败：向报表导出队列发送消息失败");
            }
        } else {
            throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "创建报表导出任务失败：参数不全");
        }
    }

    /**
     * 分页查询报表导出任务列表
     */
    public Page<RPTExportTaskEntity> getRptExportTaskList(RPTExportTaskSearch search) {
        PageHelper.startPage(search.getPageNo(), search.getPageSize());
        Page<RPTExportTaskEntity> taskEntityPage = rptExportTaskMapper.getTaskList(RptCommonUtils.getSystemId(), search.getTaskCreateBy(),
                search.getReportId(), search.getReportType(), search.getBeginTaskCreateDate(), search.getEndTaskCreateDate());
        if (taskEntityPage != null && !taskEntityPage.isEmpty()) {
            for (RPTExportTaskEntity item : taskEntityPage) {
                if (item.getStatus() >= RPTReportExportStatusEnum.SUCCESS.value && StringUtils.isNotBlank(item.getReportFilePath())) {
                    item.setReportFilePath(RptExcelFileUtils.getRptExcelFileHostDir() + item.getReportFilePath());
                }

            }
        }
        return taskEntityPage;
    }


    /**
     * 往队列放报表导出任务消息体
     */
    private boolean sendRptExportTaskMessage(RPTExportTaskEntity taskEntity) {
        boolean result = false;
        if (taskEntity != null && taskEntity.getId() != null && taskEntity.getId() > 0 && RPTReportEnum.isReportId(taskEntity.getReportId())) {
            MQRPTExportTaskMessage.RPTExportTaskMessage.Builder builder = MQRPTExportTaskMessage.RPTExportTaskMessage.newBuilder();
            builder.setTaskId(taskEntity.getId())
                    .setReportId(taskEntity.getReportId())
                    .setSearchConditionJson(StringUtils.toString(taskEntity.getSearchConditionJson()))
                    .setQuarter(StringUtils.toString(taskEntity.getQuarter()));
            result = rptExportReportTaskMQSender.send(builder.build());
        }
        return result;
    }

    /**
     * 处理报表导出任务消息
     */
    public void processRptExportTaskMessage(MQRPTExportTaskMessage.RPTExportTaskMessage message) {
        RPTExportTaskEntity taskEntity = rptExportTaskMapper.get(message.getTaskId());
        RPTReportEnum reportEnum = RPTReportEnum.valueOf(message.getReportId());
        SXSSFWorkbook xBook = null;
        String fileName = null;
        if (taskEntity != null && reportEnum != null) {
            boolean flag = false;
            taskEntity.setStatus(RPTReportExportStatusEnum.FAILURE.value);
            switch (reportEnum) {
                case CT_CUSTOMER_NEW_ORDER_DAILY_RPT:
                    xBook = ctCustomerOrderDailyRptService.exportCustomerNewOrderDailyRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_NEW_ORDER_DAILY_RPT:
                    xBook = customerOrderDailyRptService.exportCustomerNewOrderDailyRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SPECIAL_CHARGE_AREA_RPT_1:
                    xBook = specialChargeAreaRptService.exportSpecialChargeCityRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SPECIAL_CHARGE_AREA_RPT_2:
                    xBook = specialChargeAreaRptService.exportSpecialChargeByCounty(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case TRAVEL_CHARGE_RANK_RPT:
                    xBook = travelChargeRankRptService.exportTravelChargeRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_CHARGE_RPT:
                    xBook = ctCustomerChargeSummaryRptNewService.exportCustomerChargeRptNew(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_CHARGE_RPT:
                case FI_CUSTOMER_CHARGE_RPT:
                    xBook = customerChargeSummaryRptNewService.exportCustomerChargeRptNew(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_UNCOMPLETED_ORDER_RPT:
                    xBook = ctUncompletedOrderRptService.UncompletedOrderExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case UNCOMPLETED_ORDER_RPT:
                    xBook = uncompletedOrderRptService.UncompletedOrderExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CANCELLED_ORDER_RPT:
                    xBook = cancelledOrderRptService.cancelledOrderRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_COMPLETED_DAILY_RPT:
                    xBook = gradedOrderRptService.exportKefuCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case PROVINCE_COMPLETED_ORDER_RPT:
                    xBook = gradedOrderRptService.exportProvinceCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CITY_COMPLETED_ORDER_RPT:
                    xBook = gradedOrderRptService.exportCityCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case COUNTY_COMPLETED_ORDER_RPT:
                    xBook = gradedOrderRptService.exportCountyCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case ORDER_SERVICE_POINT_FEE_RPT:
                    xBook = gradedOrderRptService.exportOrderServicePointFeeRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case DEVELOP_AVERAGE_ORDER_FEE:
                    xBook = gradedOrderRptService.exportDevelopAverageFeeRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_REMINDER_RPT:
                    xBook = customerReminderRptService.customerReminderExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case REMINDER_RESPONSETIME_RPT:
                    xBook = reminderResponseTimeRptService.reminderResponseTimeExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case COMPLETED_ORDER_RPT:
                case FI_COMPLETED_ORDER_RPT:
                    xBook = completedOrderDetailsService.completedOrderDetailsExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case AREA_PLAN_ORDER_RPT:
                    xBook = areaOrderPlanDailyRptService.areaOrderPlanDayRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_PLAN_ORDER_RPT:
                    xBook = keFuOrderPlanDailyRptService.keFuOrderPlanDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_PLAN_ORDER_RPT:
                case FI_CUSTOMER_PLAN_ORDER_RPT:
                    xBook = customerOrderPlanDailyRptService.customerOrderPlanDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_PLAN_ORDER_RPT:
                    xBook = ctCustomerOrderPlanDailyRptService.customerOrderPlanDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_CANCELLED_ORDER_RPT:
                    xBook = keFuOrderCancelledDailyRptService.keFuOrderCancelledDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_CHARGE_RPT:
                case SP_SERVICEPOINT_CHARGE_RPT:
                case FI_SERVICEPOINT_CHARGE_RPT:
                    xBook = rptServicePonintWriteService.PointWriteOffExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case ORDER_SOURCE_RPT:
                    xBook = rptOrderDailyWorkService.OrderDailyWorkExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case EXPLOIT_DETAIL_RPT:
                    xBook = exploitDetailRptService.exploitDetailExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_COMPLETE_TIME_RPT:
                    xBook = keFuCompleteTimeRptService.keFuCompleteTimeRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSOTMER_ORDER_TIME_RPT:
                case CT_CUSOTMER_ORDER_TIME_RPT:
                    xBook = customerOrderTimeRptService.customerOrderTimeRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_MONTH_ORDER_RPT:
                case FI_CUSTOMER_MONTH_ORDER_RPT:
                    xBook = customerMonthPlanDailyRptService.customerOrderMonthRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case DISPATCH_ORDER_RPT:
                    xBook = dispatchListInformationRptService.DisPatchListInfoRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_COMPLETED_MONTH_ORDER_RPT:
                    xBook = keFuCompletedMonthRptService.keFuCompletedMonthRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SMS_QTY_STATISTICS_RPT:
                    xBook = smsQtyStatisticsService.messageNumRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case COMPLAIN_STATISTICS_DAILY_RPT:
                    xBook = complainStatisticsDailyrPptService.ComplainStatisticsDailyExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_INVOICE_RPT:
                case FI_SERVICEPOINT_INVOICE_RPT:
                    xBook = servicePointInvoiceRptService.servicePointInvoiceRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_FINANCE_RPT:
                case FI_CUSTOMER_FINANCE_RPT:
                    xBook = customerFinanceRptService.customerFinanceRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_FINANCE_RPT:
                    xBook = ctCustomerFinanceRptService.customerFinanceRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_ACCOUNT_BALANCE_RPT:
                case FI_SERVICEPOINT_ACCOUNT_BALANCE_RPT:
                    xBook = servicePointBalanceRptService.servicePointBalanceRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case GRADE_QTY_DAILY_RPT:
                    xBook = gradeQtyDailyRptService.gradeQtyDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_COMPLAIN_RPT:
                    xBook = customerComplainRptService.customerComplainExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SALES_PERFORMANCE_RPT:
                    xBook = customerPerformanceRptService.salesPerformanceMonthRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SALES_CUSTOMER_PERFORMANCE_RPT:
                    xBook = customerPerformanceRptService.customerPerformanceMonthRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_RECEIVABLE_SUMMARY_RPT:
                case FI_CUSTOMER_RECEIVABLE_SUMMARY_RPT:
                    xBook = customerReceivableSummaryRptService.customerReceivableRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_PAY_SUMMARY_RPT:
                case FI_SERVICEPOINT_PAY_SUMMARY_RPT:
                    xBook = servicePointChargeRptService.servicePointPaySummaryExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_COST_PER_RPT:
                    xBook = servicePointChargeRptService.servicePointCostPerExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_REVENUE_RPT:
                case FI_CUSTOMER_REVENUE_RPT:
                    xBook = customerRevenueRptService.customerRevenueExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case RECHARGE_RECORD_RPT:
                case FI_RECHARGE_RECORD_RPT:
                    xBook = rechargeRecordRptService.rechargeRecordRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CHARGE_DAILY_RPT:
                    xBook = chargeDailyRptService.chargeDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICE_POINT_COVERAGE_RPT:
                    xBook = servicePointCoverageRptService.servicePointCoverAreasRptExport(taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICE_POINT_NOCOVERAGE_RPT:
                    xBook = servicePointCoverageRptService.servicePointNoCoverAreasRptExport(taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_PAYMENT_SUMMARY_RPT:
                case FI_SERVICEPOINT_PAYMENT_SUMMARY_RPT:
                    xBook = servicePointPaymentSummaryRptService.servicePointPaymentSummaryRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_RECHARGE_SUMMARY_RPT:
                case FI_CUSTOMER_RECHARGE_SUMMARY_RPT:
                    xBook = customerRechargeSummaryRptService.customerRechargeSummaryRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case PROVINCE_ORDER_COMPLAIN_RPT:
                    xBook = complainRatioDailyRptService.exportProvinceComplainCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CITY_ORDER_COMPLAIN_RPT:
                    xBook = complainRatioDailyRptService.exportCityComplainCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case COUNTY_ORDER_COMPLAIN_RPT:
                    xBook = complainRatioDailyRptService.exportCountyComplainCompletedRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_BASE_INFO_RPT:
                    xBook = servicePointBaseInfoRptService.servicePointBaseInfoRptExportNew(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case RPT_ORDER_AUDIT_ABNORMAL:
                case FI_RPT_ORDER_AUDIT_ABNORMAL:
                    xBook = abnormalFinancialReviewRptService.abnormalPlanDailyRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_PRAISE_DETAILS_RPT:
                    xBook = keFuPraiseDetailsRptService.keFuPraiseDetailsRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_PRAISE_DETAILS_RPT:
                    xBook = customerPraiseDetailsRptService.customerPraiseDetailsRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_PRAISE_DETAILS_RPT:
                    xBook = ctCustomerPraiseDetailsRptService.customerPraiseDetailsRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_PRAISE_DETAILS_RPT:
                case SP_SERVICEPOINT_PRAISE_DETAILS_RPT:
                    xBook = servicePointPraiseDetailsRptService.servicePointPraiseDetailsRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_FROZEN_DAILY_RPT:
                case CT_CUSTOMER_FROZEN_DAILY_RPT:
                    xBook = customerFrozenDailyRptService.exportCustomerFrozenDailyRpt(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case FINANCIAL_REVIEW_DETAILS_RPT:
                case FI_FINANCIAL_REVIEW_DETAILS_RPT:
                    xBook = financialReviewDetailsRptService.getFinancialReviewDetailsRptExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case UNCOMPLETED_QTY_RPT:
                    xBook = uncompletedQtyRptService.unCompletedQtyExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KA_REMINDER_RESPONSETIME_RPT:
                    xBook = kaReminderResponseTimeRptService.reminderResponseTimeExport(message.getSearchConditionJson(), taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CRUSH_COVERAGE_RPT:
                    xBook = crushCoverageRptService.crushCoverAreasRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case TRAVEL_COVERAGE_RPT:
                    xBook = travelCoverageRptService.travelCoverAreasRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SERVICEPOINT_CHARGE_NEW_RPT:
                case SP_SERVICEPOINT_CHARGE_NEW_RPT:
                    xBook = rptServicePonintWriteNewService.PointWriteOffNewExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case FINANCE_COMPLETED_ORDER_RPT:
                case FI_FINANCE_COMPLETED_ORDER_RPT:
                    xBook = completedOrderNewDetailsService.completedOrderDetailsNewExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_COMPLETE_TIME_NEW_RPT:
                    xBook = keFuCompleteTimeNewRptService.keFuCompleteTimeRptNewExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_NEW_ORDER_MONTH_RPT:
                    xBook = customerOrderMonthRptService.exportCustomerNewOrderMonthRpt(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_FROZEN_MONTH_RPT:
                    xBook = customerFrozenMonthRptService.exportCustomerFrozenMonthRpt(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case KEFU_COVERAGE_RPT:
                    xBook = keFuAreaRptService.keFuAreasRptExport(taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case MASTER_APPLY_RPT:
                    xBook = masterApplyRptService.MasterApplyRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case SPECIAL_CUSTOMER_CHARGE_AREA_RPT_2:
                    xBook = customerSpecialChargeAreaRptService.exportCustomerSpecialChargeByCounty(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case RPT_AREA_CRUSH_QTY:
                    xBook = crushAreaRptService.exportCrushArea(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CANCELLED_ORDER_DAILY_RPT:
                    xBook = cancelledOrderNewRptService.cancelledOrderNewRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case RPT_KEFU_AVERAGE_ORDER_FEE:
                    xBook = keFuAverageOrderFeeRptService.keFuAverageOrderFeeRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case RPT_KAKEFU_AVERAGE_ORDER_FEE:
                    xBook = keFuAverageOrderFeeRptService.vipKeFuAverageOrderFeeRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_SIGN_RPT:
                    xBook = customerContractRptService.customerSignDetailsRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_NEW_ORDER_MONTH_RPT:
                    xBook = ctCustomerOrderMonthRptService.exportCustomerNewOrderMonthRpt(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CT_CUSTOMER_FROZEN_MONTH_RPT:
                    xBook = ctCustomerFrozenMonthRptService.exportCustomerFrozenMonthRpt(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case FI_DEPOSIT_CHARGE_SUMMARY_RPT:
                case DEPOSIT_CHARGE_SUMMARY_RPT:
                    xBook = depositRechargeSummaryRptService.depositRechargeSummaryRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case FI_DEPOSIT_CHARGE_DETAILS_RPT:
                case DEPOSIT_CHARGE_DETAILS_RPT:
                    xBook = depositRechargeSummaryRptService.depositRechargeDetailsExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
                case CUSTOMER_EVERY_DAY_COMPLETE_RPT:
                    xBook = everyDayCompleteService.areaOrderCompleteRateRptExport(message.getSearchConditionJson(),taskEntity.getReportTitle());
                    fileName = message.getTaskId() + " "+ taskEntity.getReportTitle() + RptExcelFileUtils.RPT_EXCEL_FILE_EXT_NAME;
                    break;
            }
            if (xBook != null && StringUtils.isNotBlank(fileName)) {
                Pair<Boolean, String> saveFileResponse = RptExcelFileUtils.saveFile(xBook, fileName);
                if (saveFileResponse.getValue0()) {
                    taskEntity.setStatus(RPTReportExportStatusEnum.SUCCESS.value);
                    taskEntity.setReportFilePath(saveFileResponse.getValue1());
                    flag = true;
                }
            }
            taskEntity.setTaskCompleteDate(System.currentTimeMillis());
            rptExportTaskMapper.update(taskEntity);
            if (!flag) {
                deleteRptExportTaskIdFromRedis(taskEntity.getReportId(), taskEntity.getTaskCreateBy(), taskEntity.getSearchConditionHashcode());
            }
        }
    }


    //region 检查报表是否有数据

    /**
     * 检查报表是否有数据
     */
    public boolean hasReportData(Integer reportId, String searchConditionJson) {
        boolean result = false;
        RPTReportEnum reportEnum = RPTReportEnum.valueOf(reportId);
        if (reportEnum != null) {
            switch (reportEnum) {
                case CT_CUSTOMER_NEW_ORDER_DAILY_RPT:
                    result = ctCustomerOrderDailyRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_NEW_ORDER_DAILY_RPT:
                    result = customerOrderDailyRptService.hasReportData(searchConditionJson);
                    break;
                case SPECIAL_CHARGE_AREA_RPT_1:
                    result = specialChargeAreaRptService.hasReportData(searchConditionJson);
                    break;
                case SPECIAL_CHARGE_AREA_RPT_2:
                    result = specialChargeAreaRptService.hasReportData(searchConditionJson);
                    break;
                case TRAVEL_CHARGE_RANK_RPT:
                    result = travelChargeRankRptService.hasReportData(searchConditionJson);
                    break;
                case CT_UNCOMPLETED_ORDER_RPT:
                    result = ctUncompletedOrderRptService.hasReportData(searchConditionJson);
                    break;
                case UNCOMPLETED_ORDER_RPT:
                    result = uncompletedOrderRptService.hasReportData(searchConditionJson);
                    break;
                case CANCELLED_ORDER_RPT:
                    result = cancelledOrderRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_COMPLETED_DAILY_RPT:
                    result = gradedOrderRptService.hasKefuCompletedOrderReportData(searchConditionJson);
                    break;
                case PROVINCE_COMPLETED_ORDER_RPT:
                case CITY_COMPLETED_ORDER_RPT:
                case COUNTY_COMPLETED_ORDER_RPT:
                    result = gradedOrderRptService.hasAreaCompletedOrderReportData(searchConditionJson);
                    break;
                case ORDER_SERVICE_POINT_FEE_RPT:
                    result = gradedOrderRptService.hasOrderServicePointFeeReportData(searchConditionJson);
                    break;
                case DEVELOP_AVERAGE_ORDER_FEE:
                    result = gradedOrderRptService.hasDevelopAverageFeeReportData(searchConditionJson);
                    break;
                case CUSTOMER_REMINDER_RPT:
                    result = customerReminderRptService.hasReportData(searchConditionJson);
                    break;
                case REMINDER_RESPONSETIME_RPT:
                    result = reminderResponseTimeRptService.hasReportData(searchConditionJson);
                    break;
                case COMPLETED_ORDER_RPT:
                case FI_COMPLETED_ORDER_RPT:
                    result = completedOrderDetailsService.hasReportData(searchConditionJson);
                    break;
                case AREA_PLAN_ORDER_RPT:
                    result = areaOrderPlanDailyRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_PLAN_ORDER_RPT:
                    result = keFuOrderPlanDailyRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_PLAN_ORDER_RPT:
                case FI_CUSTOMER_PLAN_ORDER_RPT:
                    result = customerOrderPlanDailyRptService.hasReportData(searchConditionJson);
                    break;
                case CT_CUSTOMER_PLAN_ORDER_RPT:
                    result = ctCustomerOrderPlanDailyRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_CANCELLED_ORDER_RPT:
                    result = keFuOrderCancelledDailyRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_CHARGE_RPT:
                case SP_SERVICEPOINT_CHARGE_RPT:
                case FI_SERVICEPOINT_CHARGE_RPT:
                    result = rptServicePonintWriteService.hasReportData(searchConditionJson);
                    break;
                case ORDER_SOURCE_RPT:
                    result = rptOrderDailyWorkService.hasReportData(searchConditionJson);
                    break;
                case EXPLOIT_DETAIL_RPT:
                    result = exploitDetailRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_COMPLETE_TIME_RPT:
                    result = keFuCompleteTimeRptService.hasReportData(searchConditionJson);
                    break;
                case CUSOTMER_ORDER_TIME_RPT:
                case CT_CUSOTMER_ORDER_TIME_RPT:
                    result = customerOrderTimeRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_MONTH_ORDER_RPT:
                case FI_CUSTOMER_MONTH_ORDER_RPT:
                    result = customerMonthPlanDailyRptService.hasReportData(searchConditionJson);
                    break;
                case DISPATCH_ORDER_RPT:
                    result = dispatchListInformationRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_COMPLETED_MONTH_ORDER_RPT:
                    result = keFuCompletedMonthRptService.hasReportData(searchConditionJson);
                    break;
                case SMS_QTY_STATISTICS_RPT:
                    result = smsQtyStatisticsService.hasReportData(searchConditionJson);
                    break;
                case COMPLAIN_STATISTICS_DAILY_RPT:
                    result = complainStatisticsDailyrPptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_REVENUE_RPT:
                case FI_CUSTOMER_REVENUE_RPT:
                    result = customerRevenueRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_INVOICE_RPT:
                case FI_SERVICEPOINT_INVOICE_RPT:
                    result = servicePointInvoiceRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_FINANCE_RPT:
                case FI_CUSTOMER_FINANCE_RPT:
                    result = customerFinanceRptService.hasReportData(searchConditionJson);
                    break;
                case CT_CUSTOMER_FINANCE_RPT:
                    result = ctCustomerFinanceRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_ACCOUNT_BALANCE_RPT:
                case FI_SERVICEPOINT_ACCOUNT_BALANCE_RPT:
                    result = servicePointBalanceRptService.hasReportData(searchConditionJson);
                    break;
                case GRADE_QTY_DAILY_RPT:
                    result = gradeQtyDailyRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_COMPLAIN_RPT:
                    result = customerComplainRptService.hasReportData(searchConditionJson);
                    break;
                case SALES_PERFORMANCE_RPT:
                    result = customerPerformanceRptService.hasReportData(searchConditionJson);
                    break;
                case SALES_CUSTOMER_PERFORMANCE_RPT:
                    result = customerPerformanceRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_RECEIVABLE_SUMMARY_RPT:
                case FI_CUSTOMER_RECEIVABLE_SUMMARY_RPT:
                    result = customerReceivableSummaryRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_PAY_SUMMARY_RPT:
                case FI_SERVICEPOINT_PAY_SUMMARY_RPT:
                    result = servicePointChargeRptService.hasServicePointPaySummaryReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_COST_PER_RPT:
                    result = servicePointChargeRptService.hasServicePointCostPerReportData(searchConditionJson);
                    break;
                case RECHARGE_RECORD_RPT:
                case FI_RECHARGE_RECORD_RPT:
                    result = rechargeRecordRptService.hasReportData(searchConditionJson);
                    break;
                case CHARGE_DAILY_RPT:
                    result = chargeDailyRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_PAYMENT_SUMMARY_RPT:
                case FI_SERVICEPOINT_PAYMENT_SUMMARY_RPT:
                    result = servicePointPaymentSummaryRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_RECHARGE_SUMMARY_RPT:
                case FI_CUSTOMER_RECHARGE_SUMMARY_RPT:
                    result =customerRechargeSummaryRptService.hasReportData(searchConditionJson);
                    break;
                case PROVINCE_ORDER_COMPLAIN_RPT:
                case CITY_ORDER_COMPLAIN_RPT:
                case COUNTY_ORDER_COMPLAIN_RPT:
                    result =complainRatioDailyRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_BASE_INFO_RPT:
                    result =servicePointBaseInfoRptService.hasReportData(searchConditionJson);
                    break;
                case RPT_ORDER_AUDIT_ABNORMAL:
                case FI_RPT_ORDER_AUDIT_ABNORMAL:
                    result =abnormalFinancialReviewRptService.hasReportData(searchConditionJson);
                    break;
                case KEFU_PRAISE_DETAILS_RPT:
                    result =keFuPraiseDetailsRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_PRAISE_DETAILS_RPT:
                    result =customerPraiseDetailsRptService.hasReportData(searchConditionJson);
                    break;
                case CT_CUSTOMER_PRAISE_DETAILS_RPT:
                    result =ctCustomerPraiseDetailsRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_PRAISE_DETAILS_RPT:
                case SP_SERVICEPOINT_PRAISE_DETAILS_RPT:
                    result =servicePointPraiseDetailsRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_FROZEN_DAILY_RPT:
                case CT_CUSTOMER_FROZEN_DAILY_RPT:
                    result =customerFrozenDailyRptService.hasReportData(searchConditionJson);
                    break;
                case FINANCIAL_REVIEW_DETAILS_RPT:
                case FI_FINANCIAL_REVIEW_DETAILS_RPT:
                    result =financialReviewDetailsRptService.hasReportData(searchConditionJson);
                    break;
                case UNCOMPLETED_QTY_RPT:
                    result =uncompletedQtyRptService.hasReportData(searchConditionJson);
                    break;
                case KA_REMINDER_RESPONSETIME_RPT:
                    result =kaReminderResponseTimeRptService.hasReportData(searchConditionJson);
                    break;
                case SERVICEPOINT_CHARGE_NEW_RPT:
                case SP_SERVICEPOINT_CHARGE_NEW_RPT:
                    result =rptServicePonintWriteNewService.hasReportData(searchConditionJson);
                    break;
                case FINANCE_COMPLETED_ORDER_RPT:
                case FI_FINANCE_COMPLETED_ORDER_RPT:
                    result =completedOrderNewDetailsService.hasReportData(searchConditionJson);
                    break;
                case KEFU_COMPLETE_TIME_NEW_RPT:
                    result =keFuCompleteTimeNewRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_NEW_ORDER_MONTH_RPT:
                    result =customerOrderMonthRptService.hasReportData(searchConditionJson);
                    break;
                case CUSTOMER_FROZEN_MONTH_RPT:
                    result =customerFrozenMonthRptService.hasReportData(searchConditionJson);
                    break;
                case MASTER_APPLY_RPT:
                    result =masterApplyRptService.hasReportData(searchConditionJson);
                    break;
                case SPECIAL_CUSTOMER_CHARGE_AREA_RPT_1:
                    result = customerSpecialChargeAreaRptService.hasReportData(searchConditionJson);
                    break;
                case SPECIAL_CUSTOMER_CHARGE_AREA_RPT_2:
                    result = customerSpecialChargeAreaRptService.hasReportData(searchConditionJson);
                    break;
                case RPT_AREA_CRUSH_QTY:
                    result = crushAreaRptService.hasReportData(searchConditionJson);
                    break;
                case CANCELLED_ORDER_DAILY_RPT:
                    result = cancelledOrderNewRptService.hasReportData(searchConditionJson);
                    break;
                case RPT_KEFU_AVERAGE_ORDER_FEE:
                    result = keFuAverageOrderFeeRptService.hasReportData(searchConditionJson);
                    break;
                case RPT_KAKEFU_AVERAGE_ORDER_FEE:
                    result = keFuAverageOrderFeeRptService.hasVipReportData(searchConditionJson);
                    break;
                case CUSTOMER_SIGN_RPT:
                    result = customerContractRptService.hasReportData(searchConditionJson);
                    break;
                case CT_CUSTOMER_NEW_ORDER_MONTH_RPT:
                    result =ctCustomerOrderMonthRptService.hasReportData(searchConditionJson);
                    break;
                case CT_CUSTOMER_FROZEN_MONTH_RPT:
                    result =ctCustomerFrozenMonthRptService.hasReportData(searchConditionJson);
                    break;
                case FI_DEPOSIT_CHARGE_SUMMARY_RPT:
                case DEPOSIT_CHARGE_SUMMARY_RPT:
                    result =depositRechargeSummaryRptService.hasReportData(searchConditionJson);
                    break;
                case FI_DEPOSIT_CHARGE_DETAILS_RPT:
                case DEPOSIT_CHARGE_DETAILS_RPT:
                    result =depositRechargeSummaryRptService.hasDepositReportData(searchConditionJson);
                    break;
                case CUSTOMER_EVERY_DAY_COMPLETE_RPT:
                    result = everyDayCompleteService.hasReportData(searchConditionJson);
                    break;
            }
        }
        return result;
    }


    //endregion 检查报表是否有数据


    //region redis

    @Autowired
    private RedisDefaultDbNewUtils redisDefaultDbNewUtils;

    private static final int RPT_EXPORT_TASK_LOCK_EXPIRED = 60;
    /**
     * reportId:systemId:taskCreateById:searchConditionJsonHashCode
     */
    private static final String RPT_EXPORT_TASK_ID = "RPT:EXPORT:TASK:%s:%s:%s:%s";

    public void saveExportTaskIdToRedis(long taskId, int reportId, long taskCreateById, int searchConditionJsonHashCode) {
        String key = String.format(RPT_EXPORT_TASK_ID, RptCommonUtils.getSystemId(), reportId, taskCreateById, searchConditionJsonHashCode);
        redisDefaultDbNewUtils.set(key, taskId, RPT_EXPORT_TASK_LOCK_EXPIRED);
    }

    public Long getRptExportTaskIdFromRedis(int reportId, long taskCreateById, int searchConditionJsonHashCode) {
        String key = String.format(RPT_EXPORT_TASK_ID, RptCommonUtils.getSystemId(), reportId, taskCreateById, searchConditionJsonHashCode);
        return redisDefaultDbNewUtils.get(key, Long.class);
    }

    public void deleteRptExportTaskIdFromRedis(int reportId, long taskCreateById, int searchConditionJsonHashCode) {
        String key = String.format(RPT_EXPORT_TASK_ID, RptCommonUtils.getSystemId(), reportId, taskCreateById, searchConditionJsonHashCode);
        redisDefaultDbNewUtils.del(key);
    }

    //endregion redis


}














