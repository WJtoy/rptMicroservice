package com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.entity.b2bcenter.md.B2BSign;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.fallback.B2BCustomerMappingFeignFallbackFactory;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.fallback.B2BCustomerPddMappingFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

/**
 * B2Bpdd微服务接口调用
 */
@FeignClient(name = "kklplus-b2b-pdd", fallbackFactory = B2BCustomerPddMappingFeignFallbackFactory.class)
public interface B2BCustomerPddMappingFeign {

    @PostMapping("/serviceSign/getList")
    MSResponse<MSPage<B2BSign>> getServiceSignList(@RequestBody B2BSign sign);
}
