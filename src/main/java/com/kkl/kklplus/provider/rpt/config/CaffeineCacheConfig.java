package com.kkl.kklplus.provider.rpt.config;

import com.github.benmanes.caffeine.cache.Cache;
import com.github.benmanes.caffeine.cache.Caffeine;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.concurrent.TimeUnit;

@Configuration
public class CaffeineCacheConfig {

    @Bean
    public Cache<String, Object> caffeineCache() {
        return Caffeine.newBuilder()
                .expireAfterWrite(4, TimeUnit.HOURS)
                .initialCapacity(100)
                .weakKeys()
                .weakValues()
                .maximumSize(10000).build();
    }

}
