package com.kkl.kklplus.provider.rpt.ms.b2bcenter.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign.B2BCustomerMappingFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class B2BCustomerMappingFeignFallbackFactory implements FallbackFactory<B2BCustomerMappingFeign> {

    @Override
    public B2BCustomerMappingFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("B2BCustomerMappingFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new B2BCustomerMappingFeign() {
            @Override
            public MSResponse<List<B2BCustomerMapping>> getAllCustomerMapping() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<String> getShopName(Long customerId, String shopId) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
