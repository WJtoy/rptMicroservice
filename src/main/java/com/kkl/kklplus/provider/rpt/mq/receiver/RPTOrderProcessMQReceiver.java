package com.kkl.kklplus.provider.rpt.mq.receiver;


import com.googlecode.protobuf.format.JsonFormat;
import com.kkl.kklplus.entity.rpt.mq.MQRPTOrderProcessMessage;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import com.kkl.kklplus.provider.rpt.service.RPTOrderProcessService;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.rabbitmq.client.Channel;
import lombok.extern.slf4j.Slf4j;
import org.springframework.amqp.core.Message;
import org.springframework.amqp.rabbit.annotation.RabbitListener;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class RPTOrderProcessMQReceiver {

    @Autowired
    private RPTOrderProcessService orderProcessService;

    @RabbitListener(queues = RPTMQConstant.MQ_RPT_ORDER_PROCESS_DELAY)
    public void onMessage(Message message, Channel channel) throws Exception {
        MQRPTOrderProcessMessage.RPTOrderProcessMessage msg = null;
        try {
            msg = MQRPTOrderProcessMessage.RPTOrderProcessMessage.parseFrom(message.getBody());
            if (msg != null) {
                orderProcessService.saveOrderProcess(msg);
            }
        } catch (Exception e) {
            if (msg != null) {
                String msgJson = new JsonFormat().printToString(msg);
                log.error("RPTOrderProcessMQReceiver.onMessage： {}，{}", msgJson, Exceptions.getStackTraceAsString(e));
            }
        } finally {
            channel.basicAck(message.getMessageProperties().getDeliveryTag(), false);
        }
    }


}
