package com.kkl.kklplus.provider.rpt.mq.receiver;


import com.googlecode.protobuf.format.JsonFormat;
import com.kkl.kklplus.entity.rpt.mq.MQRPTCreateOrderMessage;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import com.kkl.kklplus.provider.rpt.service.CreatedOrderService;
import com.rabbitmq.client.Channel;
import lombok.extern.slf4j.Slf4j;
import org.springframework.amqp.core.Message;
import org.springframework.amqp.rabbit.annotation.RabbitListener;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class RPTCreatedOrderMQReceiver {

    @RabbitListener(queues = RPTMQConstant.MQ_RPT_CREATE_ORDER)
    public void onMessage(Message message, Channel channel) throws Exception {
        MQRPTCreateOrderMessage.RPTCreateOrderMessage msg = null;
        try {
            msg = MQRPTCreateOrderMessage.RPTCreateOrderMessage.parseFrom(message.getBody());
            if (msg != null) {
                //createdOrderService.processRptCreatedOrderMessage(msg);
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
