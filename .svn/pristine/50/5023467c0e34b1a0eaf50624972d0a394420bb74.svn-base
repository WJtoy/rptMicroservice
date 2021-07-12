package com.kkl.kklplus.provider.rpt.chart.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.feign.MSServicePointStreetFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

@Component
@Slf4j
public class MSServicePointStreetFeignFallbackFactory implements FallbackFactory<MSServicePointStreetFeign> {
    @Override
    public MSServicePointStreetFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSServicePointStreetFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSServicePointStreetFeign() {

            @Override
            public MSResponse<List<NameValuePair<Integer, Long>>> findServicePointStationCountListForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Map<Long, List<NameValuePair<Integer, Long>>>> findStationCountListByProductCategoryForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<NameValuePair<Integer, Long>>> findStationCountListByAutoPlanFlagForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
