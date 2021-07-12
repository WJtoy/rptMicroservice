package com.kkl.kklplus.provider.rpt.mq.config;

import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import org.springframework.amqp.core.*;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
public class RPTOrderProcessMQConfig {



    @Bean
    public Queue rptOrderProcessQueue() {
        return new Queue(RPTMQConstant.MQ_RPT_ORDER_PROCESS_DELAY, true);
    }

    @Bean
    DirectExchange rptOrderProcessExchange() {
        return (DirectExchange) ExchangeBuilder.directExchange(RPTMQConstant.MQ_RPT_ORDER_PROCESS_DELAY).
                delayed().withArgument("x-delayed-type", "direct").build();
    }

    @Bean
    Binding bindingRPTOrderProcessExchangeMessage(Queue rptOrderProcessQueue, DirectExchange rptOrderProcessExchange) {
        return BindingBuilder.bind(rptOrderProcessQueue)
                .to(rptOrderProcessExchange)
                .with(RPTMQConstant.MQ_RPT_ORDER_PROCESS_DELAY);
    }

}
