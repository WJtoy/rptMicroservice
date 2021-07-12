package com.kkl.kklplus.provider.rpt.mq.config;

import com.kkl.kklplus.entity.b2b.mq.B2BMQConstant;
import com.kkl.kklplus.entity.rpt.mq.RPTMQConstant;
import org.springframework.amqp.core.*;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
public class RPTExportReportTaskMQConfig {

    @Bean
    public Queue rptExportReportTaskQueue() {
        return new Queue(RPTMQConstant.MQ_RPT_EXPORT_TASK, true);
    }

    @Bean
    DirectExchange rptExportReportTaskExchange() {
        return  (DirectExchange) ExchangeBuilder.directExchange(RPTMQConstant.MQ_RPT_EXPORT_TASK).
                delayed().withArgument("x-delayed-type", "direct").build();
    }

    @Bean
    Binding bindingRPTExportReportTaskExchangeMessage(Queue rptExportReportTaskQueue, DirectExchange rptExportReportTaskExchange) {
        return BindingBuilder.bind(rptExportReportTaskQueue)
                .to(rptExportReportTaskExchange)
                .with(RPTMQConstant.MQ_RPT_EXPORT_TASK);
    }

}
