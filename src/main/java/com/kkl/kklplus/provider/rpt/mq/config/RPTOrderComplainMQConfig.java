package com.kkl.kklplus.provider.rpt.mq.config;

import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import org.springframework.amqp.core.*;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
public class RPTOrderComplainMQConfig {



    @Bean
    public Queue rptOrderComplainQueue() {
        return new Queue(RPTMQConstant.MQ_RPT_UPDATE_ORDER_COMPLAIN, true);
    }

    @Bean
    DirectExchange rptOrderComplainExchange() {
        return new DirectExchange(RPTMQConstant.MQ_RPT_UPDATE_ORDER_COMPLAIN);
    }

    @Bean
    Binding bindingRPTOrderComplainExchangeMessage(Queue rptOrderComplainQueue, DirectExchange rptOrderComplainExchange) {
        return BindingBuilder.bind(rptOrderComplainQueue)
                .to(rptOrderComplainExchange)
                .with(RPTMQConstant.MQ_RPT_UPDATE_ORDER_COMPLAIN);
    }

}
