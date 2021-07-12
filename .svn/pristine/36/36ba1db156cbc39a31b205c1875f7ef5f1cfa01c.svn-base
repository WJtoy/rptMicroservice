package com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.fallback.B2BCustomerMappingFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

/**
 * B2BCenter微服务接口调用
 */
@FeignClient(name = "kklplus-b2b-center", fallbackFactory = B2BCustomerMappingFeignFallbackFactory.class)
public interface B2BCustomerMappingFeign {
    /**
     * 获取系统中所有的店铺名称
     */
    @GetMapping("/b2BCustomerMapping/getAllCustomerMapping")
    MSResponse<List<B2BCustomerMapping>> getAllCustomerMapping();

    /**
     * 获取指定客户店铺的名称
     */
    @GetMapping("/b2BCustomerMapping/getShopName")
    MSResponse<String> getShopName(@RequestParam("customerId") Long customerId, @RequestParam("shopId") String shopId);
}
