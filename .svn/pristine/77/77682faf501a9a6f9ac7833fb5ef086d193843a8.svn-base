package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDEngineer;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSEngineerFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSEngineerFeignFallbackFactory implements FallbackFactory<MSEngineerFeign> {
    @Override
    public MSEngineerFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSEngineerFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSEngineerFeign() {

            /**
             * 根据id列表获取安维列表
             *
             * @param ids
             * @return
             */
            @Override
            public MSResponse<List<MDEngineer>> findEngineersByIds(List<Long> ids, List<String> fields) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
