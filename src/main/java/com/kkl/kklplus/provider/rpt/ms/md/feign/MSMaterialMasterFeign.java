package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDActionCode;
import com.kkl.kklplus.entity.md.MDMaterial;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSErrorFeignFallbackFactory;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSMaterialMasterFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSMaterialMasterFallbackFactory.class)
public interface MSMaterialMasterFeign {

    /**
     * 根据id列表获取安配件名
     * @param ids
     * @return
     */
    @PostMapping("/material/findIdAndNameByIds")
    MSResponse<List<MDMaterial>> findMaterialById(@RequestParam("ids") List<Long> ids);


}
