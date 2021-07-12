package com.kkl.kklplus.provider.rpt.ms.md.fallback;


import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDActionCode;
import com.kkl.kklplus.entity.md.MDErrorCode;
import com.kkl.kklplus.entity.md.MDErrorType;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSErrorFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSErrorFeignFallbackFactory implements FallbackFactory<MSErrorFeign> {
    @Override
    public MSErrorFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSErrorFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSErrorFeign() {
            @Override
            public MSResponse<List<MDErrorType>> findListByErrorTypeName(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<MDErrorCode>> findListByErrorCodeName(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<MDActionCode>> findListByActionCodeName(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
