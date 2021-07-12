package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSSmsQtyFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.RequestBody;

import java.util.Map;

@Component
@Slf4j
public class MSSmsQtyFeignFallbackFactory implements FallbackFactory<MSSmsQtyFeign> {

    @Override
    public MSSmsQtyFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSSmsQtyFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSSmsQtyFeign() {

            @Override
            public MSResponse<Map<Integer, Long>> shortMessageCache(@RequestBody String date) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
