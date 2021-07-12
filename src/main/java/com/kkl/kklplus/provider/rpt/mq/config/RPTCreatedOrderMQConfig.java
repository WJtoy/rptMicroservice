package com.kkl.kklplus.provider.rpt.mq.config;

import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import org.springframework.amqp.core.*;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
public class RPTCreatedOrderMQConfig {

    @Bean
    public Queue rptCreatedOrderQueue() {
        return new Queue(RPTMQConstant.MQ_RPT_CREATE_ORDER, true);
    }

    @Bean
    DirectExchange rptCreatedOrderExchange() {
        return new DirectExchange(RPTMQConstant.MQ_RPT_CREATE_ORDER);
    }

    @Bean
    Binding bindingRPTCreatedOrderExchangeMessage(Queue rptCreatedOrderQueue, DirectExchange rptCreatedOrderExchange) {
        return BindingBuilder.bind(rptCreatedOrderQueue)
                .to(rptCreatedOrderExchange)
                .with(RPTMQConstant.MQ_RPT_CREATE_ORDER);
    }

}
