package com.kkl.kklplus.provider.rpt.mq.sender;

import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import lombok.extern.slf4j.Slf4j;
import om.kkl.kklplus.entity.rpt.mq.pb.MQRPTExportTaskMessage;
import org.springframework.amqp.core.MessageDeliveryMode;
import org.springframework.amqp.rabbit.core.RabbitTemplate;
import org.springframework.amqp.rabbit.support.CorrelationData;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.retry.RetryCallback;
import org.springframework.retry.support.RetryTemplate;
import org.springframework.stereotype.Component;

import java.util.concurrent.atomic.AtomicBoolean;


@Slf4j
@Component
public class RPTExportReportTaskMQSender implements RabbitTemplate.ConfirmCallback {

    private RabbitTemplate rabbitTemplate;

    private RetryTemplate retryTemplate;

    @Autowired
    public RPTExportReportTaskMQSender(RabbitTemplate kklRabbitTemplate, RetryTemplate kklRabbitRetryTemplate) {
        this.rabbitTemplate = kklRabbitTemplate;
        this.rabbitTemplate.setConfirmCallback(this);
        this.retryTemplate = kklRabbitRetryTemplate;
    }

    /**
     * 正常发送消息
     *
     * @param message 消息体
     */
    public boolean send(MQRPTExportTaskMessage.RPTExportTaskMessage message) {
        AtomicBoolean result = new AtomicBoolean(false);
        try {
            retryTemplate.execute((RetryCallback<Object, Exception>) context -> {
                context.setAttribute(RPTMQConstant.RETRY_CONTEXT_ATTRIBUTE_KEY_MESSAGE, message);
                rabbitTemplate.convertAndSend(
                        RPTMQConstant.MQ_RPT_EXPORT_TASK,
                        RPTMQConstant.MQ_RPT_EXPORT_TASK,
                        message.toByteArray(),
                        msg -> {
                            msg.getMessageProperties().setDelay(10 * 1000);
                            msg.getMessageProperties().setDeliveryMode(MessageDeliveryMode.PERSISTENT);
                            return msg;
                        },
                        new CorrelationData());
                result.set(true);
                return null;
            }, context -> {
                Object msgObj = context.getAttribute(RPTMQConstant.RETRY_CONTEXT_ATTRIBUTE_KEY_MESSAGE);
                MQRPTExportTaskMessage.RPTExportTaskMessage msg = MQRPTExportTaskMessage.RPTExportTaskMessage.parseFrom((byte[]) msgObj);
                Throwable throwable = context.getLastThrowable();
                log.error("normal send error {}, {}", throwable.getLocalizedMessage(), msg);
                return null;
            });
        } catch (Exception e) {
            log.error("RPTExportReportTaskMQSender.send：{}", e.getLocalizedMessage());
        }
        return result.get();
    }

    @Override
    public void confirm(CorrelationData correlationData, boolean ack, String cause) {

    }
}
