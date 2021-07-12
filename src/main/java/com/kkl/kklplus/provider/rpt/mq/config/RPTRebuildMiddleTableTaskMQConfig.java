package com.kkl.kklplus.provider.rpt.mq.config;

import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import org.springframework.amqp.core.*;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
public class RPTRebuildMiddleTableTaskMQConfig {

    @Bean
    public Queue rptRebuildMiddleTableTaskQueue() {
        return new Queue(RPTMQConstant.MQ_RPT_REBUILD_MIDDLE_TABLE_TASK, true);
    }

    @Bean
    DirectExchange rptRebuildMiddleTableTaskExchange() {
        return (DirectExchange) ExchangeBuilder.directExchange(RPTMQConstant.MQ_RPT_REBUILD_MIDDLE_TABLE_TASK).
                delayed().withArgument("x-delayed-type", "direct").build();
    }

    @Bean
    Binding bindingRPTRebuildMiddleTableTaskExchangeMessage(Queue rptRebuildMiddleTableTaskQueue, DirectExchange rptRebuildMiddleTableTaskExchange) {
        return BindingBuilder.bind(rptRebuildMiddleTableTaskQueue)
                .to(rptRebuildMiddleTableTaskExchange)
                .with(RPTMQConstant.MQ_RPT_REBUILD_MIDDLE_TABLE_TASK);
    }

}
