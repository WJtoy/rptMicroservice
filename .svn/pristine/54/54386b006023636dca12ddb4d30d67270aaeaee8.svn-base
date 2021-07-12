package com.kkl.kklplus.provider.rpt.ms.sys.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSAreaFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSAreaFeignFallbackFactory  implements FallbackFactory<MSAreaFeign> {
    @Override
    public MSAreaFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSAreaFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSAreaFeign()
        {
            @Override
            public MSResponse<MSPage<Long>> getCrushList(int pageNo, int pageSize, List<Long> productCategoryIds) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<Long>> getTravelList(int pageNo, int pageSize, List<Long> productCategoryIds) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

        };
    }
}
