package com.kkl.kklplus.provider.rpt.mq.receiver;


import com.googlecode.protobuf.format.JsonFormat;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import com.kkl.kklplus.provider.rpt.service.RptExportTaskService;
import com.rabbitmq.client.Channel;
import lombok.extern.slf4j.Slf4j;
import om.kkl.kklplus.entity.rpt.mq.pb.MQRPTExportTaskMessage;
import org.springframework.amqp.core.Message;
import org.springframework.amqp.rabbit.annotation.RabbitListener;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class RPTExportReportTaskMQReceiver {

    @Autowired
    private RptExportTaskService rptExportTaskService;

    @RabbitListener(queues = RPTMQConstant.MQ_RPT_EXPORT_TASK)
    public void onMessage(Message message, Channel channel) throws Exception {
        MQRPTExportTaskMessage.RPTExportTaskMessage msg = null;
        try {
            msg = MQRPTExportTaskMessage.RPTExportTaskMessage.parseFrom(message.getBody());
            if (msg != null && msg.getTaskId() > 0) {
                rptExportTaskService.processRptExportTaskMessage(msg);
            }
        } catch (Exception e) {
            if (msg != null) {
                String msgJson = new JsonFormat().printToString(msg);
                log.error("RPTExportReportTaskMQReceiver.onMessage： {}，{}", msgJson, e);
            }
        } finally {
            channel.basicAck(message.getMessageProperties().getDeliveryTag(), false);
        }
    }

}
