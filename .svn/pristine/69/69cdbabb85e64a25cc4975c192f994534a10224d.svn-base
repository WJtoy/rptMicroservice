package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDServiceType;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSServiceTypeFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSServiceTypeFeignFallbackFactory.class)
public interface MSServiceTypeFeign {

    @GetMapping("/serviceType/findAllList")
    MSResponse<List<MDServiceType>> findAllList();
}
