package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDServiceType;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSServiceTypeFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSServiceTypeFeignFallbackFactory implements FallbackFactory<MSServiceTypeFeign> {
    @Override
    public MSServiceTypeFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSServiceTypeFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSServiceTypeFeign() {
            @Override
            public MSResponse<List<MDServiceType>> findAllList() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
