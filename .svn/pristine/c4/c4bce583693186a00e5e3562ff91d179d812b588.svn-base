package com.kkl.kklplus.provider.rpt.ms.sys.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.provider.rpt.ms.sys.fallback.MSAreaFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;

import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSAreaFeignFallbackFactory.class)
public interface MSAreaFeign {

    @PostMapping("regionPermission/findCrushSubAreaList")
    MSResponse<MSPage<Long>> getCrushList(@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize,@RequestBody  List<Long> productCategoryIds);

    @PostMapping("regionPermission/findRemotefeeSubAreaList")
    MSResponse<MSPage<Long>> getTravelList(@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize, @RequestBody List<Long> productCategoryIds);
}
