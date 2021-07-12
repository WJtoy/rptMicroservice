package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.md.MDProduct;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSProductFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;


@Component
@Slf4j
public class MSProductFeignFallbackFactory implements FallbackFactory<MSProductFeign> {
    @Override
    public MSProductFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSProductFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSProductFeign() {

            @Override
            public MSResponse<List<MDProduct>> findListByConditions(MDProduct mdProduct) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<NameValuePair<Long, String>>> findListByProductIdsForRPT(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
