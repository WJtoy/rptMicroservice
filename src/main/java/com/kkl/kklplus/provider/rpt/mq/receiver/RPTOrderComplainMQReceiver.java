package com.kkl.kklplus.provider.rpt.mq.receiver;


import com.googlecode.protobuf.format.JsonFormat;
import com.kkl.kklplus.entity.rpt.mq.MQRPTUpdateOrderComplainMessage;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import com.kkl.kklplus.provider.rpt.service.ComplainRatioDailyRptService;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.rabbitmq.client.Channel;
import lombok.extern.slf4j.Slf4j;
import org.springframework.amqp.core.Message;
import org.springframework.amqp.rabbit.annotation.RabbitListener;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class RPTOrderComplainMQReceiver {

    @Autowired
    private ComplainRatioDailyRptService complainRatioDailyRptService;

    @RabbitListener(queues = RPTMQConstant.MQ_RPT_UPDATE_ORDER_COMPLAIN)
    public void onMessage(Message message, Channel channel) throws Exception {
        MQRPTUpdateOrderComplainMessage.MQOrderComplainMessage msg = null;
        try {
            msg = MQRPTUpdateOrderComplainMessage.MQOrderComplainMessage.parseFrom(message.getBody());
            if (msg != null) {
                complainRatioDailyRptService.saveOrderComplainMQ(msg);
            }
        } catch (Exception e) {
            if (msg != null) {
                String msgJson = new JsonFormat().printToString(msg);
                log.error("RPTOrderComplainMQReceiver.onMessage： {}，{}", msgJson, Exceptions.getStackTraceAsString(e));
            }
        } finally {
            channel.basicAck(message.getMessageProperties().getDeliveryTag(), false);
        }
    }


}
