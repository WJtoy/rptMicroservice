package com.kkl.kklplus.provider.rpt.config;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.kkl.kklplus.starter.redis.config.RedisConstant;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

@Component
public class AdapterConfig {
    @Bean(RedisConstant.BEAN_NAME_REDIS_GSON)
    public Gson gsonBean() {
        return new GsonBuilder()
//                .registerTypeAdapter(MDCustomer.class, com.kkl.kklplus.provider.rpt.utils.CustomerSimpleAdapter.getInstance())
                .create();
    }
}
