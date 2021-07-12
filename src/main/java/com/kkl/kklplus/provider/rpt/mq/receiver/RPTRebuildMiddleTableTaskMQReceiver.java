package com.kkl.kklplus.provider.rpt.mq.receiver;


import com.googlecode.protobuf.format.JsonFormat;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import com.kkl.kklplus.provider.rpt.service.RebuildMiddleTableService;
import com.rabbitmq.client.Channel;
import lombok.extern.slf4j.Slf4j;
import om.kkl.kklplus.entity.rpt.mq.pb.MQRPTRebuildMiddleTableTaskMessage;
import org.springframework.amqp.core.Message;
import org.springframework.amqp.rabbit.annotation.RabbitListener;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class RPTRebuildMiddleTableTaskMQReceiver {

    @Autowired
    private RebuildMiddleTableService rebuildMiddleTableService;

    @RabbitListener(queues = RPTMQConstant.MQ_RPT_REBUILD_MIDDLE_TABLE_TASK)
    public void onMessage(Message message, Channel channel) throws Exception {
        MQRPTRebuildMiddleTableTaskMessage.RPTRebuildMiddleTableTaskMessage msg = null;
        try {
            msg = MQRPTRebuildMiddleTableTaskMessage.RPTRebuildMiddleTableTaskMessage.parseFrom(message.getBody());
            if (msg != null && msg.getTaskId() > 0) {
                rebuildMiddleTableService.processRebuildMiddleTableTaskMessage(msg);
            }
        } catch (Exception e) {
            if (msg != null) {
                String msgJson = new JsonFormat().printToString(msg);
                log.error("RPTRebuildMiddleTableTaskMQReceiver.onMessage：{}，{}", msgJson, e);
            }
        } finally {
            channel.basicAck(message.getMessageProperties().getDeliveryTag(), false);
        }
    }

}
