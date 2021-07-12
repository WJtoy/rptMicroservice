package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.md.MDProduct;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSProductFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSProductFeignFallbackFactory.class)
public interface MSProductFeign {

    /**
     * 按条件返回产品信息列表
     */
    @PostMapping("/product/findListByConditions")
    MSResponse<List<MDProduct>> findListByConditions(@RequestBody MDProduct mdProduct);

    /**
     * 按条件返回产品信息列表
     */
    @PostMapping("/productVerSecond/findListByProductIdsFromCacheForRPT")
    MSResponse<List<NameValuePair<Long, String>>> findListByProductIdsForRPT(@RequestBody List<Long> ids);
}
