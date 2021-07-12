package com.kkl.kklplus.provider.rpt.service;

import com.kkl.kklplus.entity.rpt.common.RPTOrderProcessTypeEnum;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class RPTOrderProcessService {

    @Autowired
    private CancelledOrderDailyRptService cancelledOrderDailyRptService;

    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    @Autowired
    private CreatedOrderService createdOrderService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    public void saveOrderProcess(MQRPTOrderProcessMessage.RPTOrderProcessMessage msg){
       if(msg.getProcessType() == RPTOrderProcessTypeEnum.CANCEL.getValue() || msg.getProcessType() == RPTOrderProcessTypeEnum.RETURN.getValue()){
          cancelledOrderDailyRptService.save(msg);
           cancelledOrderRptService.save(msg);
       }
       if (msg.getProcessType()==RPTOrderProcessTypeEnum.GRADE.getValue()){
           gradedOrderRptService.saveGradeOrderOfMQToRptDB(msg);
       }

       if(msg.getProcessType() == RPTOrderProcessTypeEnum.CREATE.getValue()){
           createdOrderService.processRptCreatedOrderMessage(msg);
       }

    }
}
