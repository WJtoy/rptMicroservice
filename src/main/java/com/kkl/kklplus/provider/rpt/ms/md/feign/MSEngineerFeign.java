package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDEngineer;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSEngineerFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSEngineerFeignFallbackFactory.class)
public interface MSEngineerFeign {

    /**
     * 根据id列表获取安维列表
     * @param ids
     * @return
     */
    @PostMapping("/engineer/findEngineersByIds")
    MSResponse<List<MDEngineer>>  findEngineersByIds(@RequestBody List<Long> ids, @RequestParam("fields") List<String> fields);
}
