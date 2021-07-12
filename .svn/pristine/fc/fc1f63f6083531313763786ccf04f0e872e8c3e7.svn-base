package com.kkl.kklplus.provider.rpt.chart.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.feign.MSServicePointQtyFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

@Component
@Slf4j
public class MSServicePointQtyFeignFallbackFactory implements FallbackFactory<MSServicePointQtyFeign> {


    @Override
    public MSServicePointQtyFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSServicePointQtyFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSServicePointQtyFeign() {

            @Override
            public MSResponse<List<NameValuePair<Integer, Long>>> findServicePointCountListByDegreeForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Map<Long, List<NameValuePair<Integer, Integer>>>> findServicePointCountListByCategoryForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<NameValuePair<Long, Long>>> findServicePointCountListByCategoryAndAutoPlanFlagForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
