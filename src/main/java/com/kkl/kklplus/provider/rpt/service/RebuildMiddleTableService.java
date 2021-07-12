package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.kkl.kklplus.entity.rpt.RPTRebuildMiddleTableTaskEntity;
import com.kkl.kklplus.entity.rpt.common.*;
import com.kkl.kklplus.entity.rpt.exception.RPTBaseException;
import com.kkl.kklplus.entity.rpt.search.RPTRebuildMiddleTableTaskSearch;
import com.kkl.kklplus.provider.rpt.chart.service.OrderCrushQtyChartService;
import com.kkl.kklplus.provider.rpt.chart.service.OrderQtyDailyChartService;
import com.kkl.kklplus.provider.rpt.chart.service.ServicePointQtyStatisticsService;
import com.kkl.kklplus.provider.rpt.chart.service.ServicePointStreetQtyService;
import com.kkl.kklplus.provider.rpt.mapper.RebuildMiddleTableTaskMapper;
import com.kkl.kklplus.provider.rpt.mq.receiver.RPTRebuildMiddleTableTaskMQReceiver;
import com.kkl.kklplus.provider.rpt.mq.sender.RPTRebuildMiddleTableTaskMQSender;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.starter.redis.utils.RedisDefaultDbNewUtils;
import lombok.extern.slf4j.Slf4j;
import om.kkl.kklplus.entity.rpt.mq.pb.MQRPTRebuildMiddleTableTaskMessage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.Date;

/**
 * 重建中间表数据服务
 */
@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class RebuildMiddleTableService {

    @Autowired
    RPTRebuildMiddleTableTaskMQSender rptRebuildMiddleTableTaskMQSender;
    @Autowired
    RPTRebuildMiddleTableTaskMQReceiver rptRebuildMiddleTableTaskMQReceiver;
    @Resource
    RebuildMiddleTableTaskMapper rebuildMiddleTableTaskMapper;

    @Autowired
    private SpecialChargeAreaRptService specialChargeAreaRptService;
    @Autowired
    private TravelChargeRankRptService travelChargeRankRptService;
    @Autowired
    private CompletedOrderRptService completedOrderRptService;
    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;
    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;
    @Autowired
    private CustomerChargeSummaryRptService customerChargeSummaryRptService;
    @Autowired
    private ServicePointCompletedOrderRptService servicePointCompletedOrderRptService;
    @Autowired
    private ServicePointWriteOffRptService servicePointWriteOffRptService;
    @Autowired
    private CustomerReminderRptService customerReminderRptService;

    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    @Autowired
    private CreatedOrderService createdOrderService;

    @Autowired
    private CancelledOrderDailyRptService cancelledOrderDailyRptService;

    @Autowired
    private CustomerSalesMappingService customerSalesMappingService;

    @Autowired
    private KeFuCompleteTimeRptService keFuCompleteTimeRptService;

    @Autowired
    private CustomerOrderTimeRptService customerOrderTimeRptService;

    @Autowired
    private SMSQtyStatisticsService smsQtyStatisticsService;

    @Autowired
    private CustomerRevenueRptService customerRevenueRptService;

    @Autowired
    private GradeQtyDailyRptService gradeQtyDailyRptService;

    @Autowired
    private CustomerPerformanceRptService customerPerformanceRptService;

    @Autowired
    private CustomerChargeSummaryRptNewService customerChargeSummaryRptNewService;

    @Autowired
    private ServicePointChargeRptService servicePointChargeRptService;
    @Autowired
    private ServicePointChargeRptNewService servicePointChargeRptNewService;

    @Autowired
    private ComplainRatioDailyRptService complainRatioDailyRptService;

    @Autowired
    private AbnormalFinancialReviewRptService abnormalFinancialReviewRptService;

    @Autowired
    private ServicePointQtyStatisticsService servicePointQtyStatisticsService;

    @Autowired
    private ServicePointStreetQtyService servicePointStreetQtyService;

    @Autowired
    private OrderQtyDailyChartService orderQtyDailyChartService;

    @Autowired
    private OrderCrushQtyChartService orderCrushQtyChartService;

    @Autowired
    private CustomerSpecialChargeAreaRptService customerSpecialChargeAreaRptService;


    @Autowired
    private CrushAreaRptService crushAreaRptService;

    @Autowired
    private CustomerInfoMappingService customerInfoMappingService;

    @Autowired
    private SysUserRegionMappingService sysUserRegionMappingService;

    @Autowired
    private SysUserCustomerMappingService sysUserCustomerMappingService;
    /**
     * 创建重建中间表的任务
     */
    @Transactional()
    public void createRebuildMiddleTableTask(RPTRebuildMiddleTableTaskEntity params) {
        checkRebuildMiddleTableTask(params);
        params.setSystemId(RptCommonUtils.getSystemId());
        params.setTaskCreateDate(System.currentTimeMillis());
        params.setStatus(RPTRebuildStatusEnum.START.value);
        rebuildMiddleTableTaskMapper.insert(params);
        MQRPTRebuildMiddleTableTaskMessage.RPTRebuildMiddleTableTaskMessage.Builder builder = MQRPTRebuildMiddleTableTaskMessage.RPTRebuildMiddleTableTaskMessage.newBuilder();
        builder.setTaskId(params.getId());
        boolean result = rptRebuildMiddleTableTaskMQSender.send(builder.build());
        if (result) {
            saveRebuildMiddleTableTaskIdToRedis(params.getId(), params.getMiddleTableId());
        } else {
            throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：向队列发送消息失败");
        }
    }


    private void checkRebuildMiddleTableTask(RPTRebuildMiddleTableTaskEntity params) {
        if (params != null && RPTMiddleTableEnum.isMiddleTableId(params.getMiddleTableId())
                && RPTMiddleTableTypeEnum.isMiddleTableType(params.getMiddleTableType())
                && RPTRebuildOperationTypeEnum.isOperationType(params.getOperationType())) {
            RPTMiddleTableTypeEnum tableTypeEnum = RPTMiddleTableTypeEnum.valueOf(params.getMiddleTableType());
            if (tableTypeEnum == RPTMiddleTableTypeEnum.DAY) {
                if (params.getBeginDate() == null || params.getBeginDate() == 0
                        || params.getEndDate() == null || params.getEndDate() == 0) {
                    throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：参数不全");
                }
            } else if (tableTypeEnum == RPTMiddleTableTypeEnum.YEAR_MONTH) {
                if (params.getSelectedYear() == null || params.getSelectedYear() == 0
                        || params.getSelectedMonth() == null || params.getSelectedMonth() == 0) {
                    throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：参数不全");
                }
            } else {
                throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：参数不全");
            }

            Long taskId = getRebuildMiddleTableTaskIdFromRedis(params.getMiddleTableId());
            if (taskId != null) {
                throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：该中间表正在重建中");
            }
        } else {
            throw new RPTBaseException(RPTErrorCode.RPT_OPERATE_FAILURE, "重建中间表数据失败：参数不全");
        }
    }


    /**
     * 分页查询任务列表
     */
    public Page<RPTRebuildMiddleTableTaskEntity> getRebuildMiddleTableTaskList(RPTRebuildMiddleTableTaskSearch search) {
        PageHelper.startPage(search.getPageNo(), search.getPageSize());
        return rebuildMiddleTableTaskMapper.getTaskList(RptCommonUtils.getSystemId(), search.getMiddleTableId());
    }


    /**
     * 处理报表中间表重建任务消息
     */
    public void processRebuildMiddleTableTaskMessage(MQRPTRebuildMiddleTableTaskMessage.RPTRebuildMiddleTableTaskMessage message) {
        RPTRebuildMiddleTableTaskEntity taskEntity = rebuildMiddleTableTaskMapper.get(message.getTaskId());
        if (taskEntity != null) {
            RPTMiddleTableEnum tableEnum = RPTMiddleTableEnum.valueOf(taskEntity.getMiddleTableId());
            RPTRebuildOperationTypeEnum operationTypeEnum = RPTRebuildOperationTypeEnum.valueOf(taskEntity.getOperationType());
            boolean flag = false;
            taskEntity.setStatus(RPTRebuildStatusEnum.FAILURE.value);
            Date now = new Date();
            int hourIndex = DateUtils.getHourOfDay(now);
            long intervalTime = taskEntity.getTaskCreateDate() - now.getTime();
           //TODO: 仅允许晚上8:00 ~ 8:00执行写中间表的任务
//            if ((hourIndex >= 13 || hourIndex < 8) && intervalTime < 5 * 60 * 1000) {
          if (intervalTime < 5 * 60 * 1000) {
                switch (tableEnum) {
                    case RPT_AREA_SPECIAL_CHARGE:
                        flag = specialChargeAreaRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getSelectedYear(), taskEntity.getSelectedMonth());
                        break;
                    case RPT_TRAVEL_CHARGE_RANK:
                        flag = travelChargeRankRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getSelectedYear(), taskEntity.getSelectedMonth());
                        break;
                    case RPT_COMPLETED_ORDER:
                        flag = completedOrderRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_CANCELLED_ORDER:
                        flag = cancelledOrderRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_WRITE_OFF:
                        flag = customerWriteOffRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_FINANCE_MONTHLY:
                    case RPT_CUSTOMER_ORDER_QTY_MONTHLY:
                        flag = customerChargeSummaryRptNewService.rebuildMiddleTableData(tableEnum, operationTypeEnum, taskEntity.getSelectedYear(), taskEntity.getSelectedMonth());
                        break;
                    case RPT_GRADED_ORDER:
                        flag = gradedOrderRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;

                    case RPT_SERVICEPOINT_COMPLETED_ORDER:
                        flag = servicePointCompletedOrderRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_SERVICEPOINT_WRITE_OFF:
                        flag = servicePointWriteOffRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_REMINDER_DAILY:
                        flag = customerReminderRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_CREATED_ORDER:
                        flag = createdOrderService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_CANCELLED_ORDER_DAILY:
                        flag = cancelledOrderDailyRptService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_MD_CUSTOMER_SALES:
                        flag = customerSalesMappingService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_ORDER_TIME:
                        flag = customerOrderTimeRptService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SMS_QTY_STATISTICS:
                        flag = smsQtyStatisticsService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_REVENUE:
                        flag = customerRevenueRptService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_GRADE_QTY_DAILY:
                        flag = gradeQtyDailyRptService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_PERFORMANCE_MONTHLY:
                        flag = customerPerformanceRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getBeginDate(), taskEntity.getEndDate());
                        break;
                    case RPT_SERVICEPOINT_CHARGE:
//                        flag = servicePointChargeRptService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        flag = servicePointChargeRptNewService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_KEFU_COMPLETE_TIME_CREATE:
                        flag = keFuCompleteTimeRptService.rebuildMiddleTableCreateOrderData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_KEFU_COMPLETE_TIME_CLOSE:
                        flag = keFuCompleteTimeRptService.rebuildMiddleTableCloseData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_KEFU_COMPLETE_TIME_PLAN:
                        flag = keFuCompleteTimeRptService.rebuildMiddleTablePlanTypeData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_KEFU_COMPLETE_TIME_COMPLAIN:
                        flag = keFuCompleteTimeRptService.rebuildMiddleTableComplainData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_ORDER_COMPLAIN:
                        flag = complainRatioDailyRptService.rebuildMiddleTableComplainData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_ORDER_AUDIT_ABNORMAL:
                        flag = abnormalFinancialReviewRptService.rebuildMiddleAbnormalTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SERVICEPOINT_QTY:
                        flag = servicePointQtyStatisticsService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SERVICEPOINT_STREET:
                        flag = servicePointStreetQtyService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_ORDERCRUSH_QTY:
                        flag = orderCrushQtyChartService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_ORDERDAILY_QTY:
                        flag = orderQtyDailyChartService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SERVICE_POINT_CHARGE_PLATFORMFEE:
                        flag = servicePointChargeRptService.rebuildUpdatePlatFromFee(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SERVICE_POINT_CHARGE_QUARTER:
                        flag = servicePointChargeRptService.rebuildUpdateQuarterMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SERVICE_POINT_CHARGE_PLATFORMFEENEW:
                        flag = servicePointChargeRptService.rebuildUpdatePlatFromFeeNew(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_CUSTOMER_AREA_SPECIAL_CHARGE:
                        flag = customerSpecialChargeAreaRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getSelectedYear(), taskEntity.getSelectedMonth());
                        break;
                    case RPT_AREA_CRUSH_QTY:
                        flag = crushAreaRptService.rebuildMiddleTableData(operationTypeEnum, taskEntity.getSelectedYear(), taskEntity.getSelectedMonth());
                        break;
                    case RPT_MD_CUSTOMER_INFO:
                        flag = customerInfoMappingService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SYS_USER_REGION:
                        flag = sysUserRegionMappingService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;
                    case RPT_SYS_USER_CUSTOMER:
                        flag = sysUserCustomerMappingService.rebuildMiddleTableData(operationTypeEnum,taskEntity.getBeginDate(),taskEntity.getEndDate());
                        break;

                }
            }
            taskEntity.setTaskCompleteDate(System.currentTimeMillis());
            if (flag) {
                taskEntity.setStatus(RPTRebuildStatusEnum.SUCCESS.value);
            }
            rebuildMiddleTableTaskMapper.update(taskEntity);
            deleteRebuildMiddleTableTaskIdFromRedis(tableEnum.getValue());
        }
    }


    //region redis

    @Autowired
    private RedisDefaultDbNewUtils redisDefaultDbNewUtils;

    private static final int RPT_REBUILD_MIDDLE_TABLE = 60 * 60 * 8;
    /**
     * systemId:middleTableId
     */
    private static final String RPT_REBUILD_MIDDLE_TABLE_TASK_ID = "RPT:REBUILD:MIDDLE:TABLE:%s:%s";

    private void saveRebuildMiddleTableTaskIdToRedis(long taskId, int middleTableId) {
        String key = String.format(RPT_REBUILD_MIDDLE_TABLE_TASK_ID, RptCommonUtils.getSystemId(), middleTableId);
        redisDefaultDbNewUtils.set(key, taskId, RPT_REBUILD_MIDDLE_TABLE);
    }

    private Long getRebuildMiddleTableTaskIdFromRedis(int middleTableId) {
        String key = String.format(RPT_REBUILD_MIDDLE_TABLE_TASK_ID, RptCommonUtils.getSystemId(), middleTableId);
        return redisDefaultDbNewUtils.get(key, Long.class);
    }

    private void deleteRebuildMiddleTableTaskIdFromRedis(int middleTableId) {
        String key = String.format(RPT_REBUILD_MIDDLE_TABLE_TASK_ID, RptCommonUtils.getSystemId(), middleTableId);
        redisDefaultDbNewUtils.del(key);
    }

    //endregion redis


}














