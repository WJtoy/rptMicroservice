package com.kkl.kklplus.provider.rpt;

import com.kkl.kklplus.provider.rpt.config.ProviderRptProperties;
import com.spring4all.swagger.EnableSwagger2Doc;
import org.springframework.amqp.rabbit.annotation.EnableRabbit;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.cloud.client.circuitbreaker.EnableCircuitBreaker;
import org.springframework.cloud.netflix.eureka.EnableEurekaClient;
import org.springframework.cloud.netflix.feign.EnableFeignClients;

@SpringBootApplication
@EnableEurekaClient
@EnableSwagger2Doc
@EnableCircuitBreaker
@EnableFeignClients
@EnableRabbit
@EnableConfigurationProperties(ProviderRptProperties.class)
public class ProviderRPTApplication {

    public static void main(String[] args) {
        SpringApplication.run(ProviderRPTApplication.class, args);
    }

}
