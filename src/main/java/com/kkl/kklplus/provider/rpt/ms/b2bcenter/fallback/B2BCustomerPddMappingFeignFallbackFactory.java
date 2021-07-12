package com.kkl.kklplus.provider.rpt.ms.b2bcenter.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.entity.b2bcenter.md.B2BSign;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign.B2BCustomerMappingFeign;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign.B2BCustomerPddMappingFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class B2BCustomerPddMappingFeignFallbackFactory implements FallbackFactory<B2BCustomerPddMappingFeign> {

    @Override
    public B2BCustomerPddMappingFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("B2BCustomerPddMappingFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new B2BCustomerPddMappingFeign() {


            @Override
            public MSResponse<MSPage<B2BSign>> getServiceSignList(B2BSign sign) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
