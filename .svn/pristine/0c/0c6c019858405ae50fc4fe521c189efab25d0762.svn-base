package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDActionCode;
import com.kkl.kklplus.entity.md.MDErrorCode;
import com.kkl.kklplus.entity.md.MDErrorType;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSErrorFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSErrorFeignFallbackFactory.class)
public interface MSErrorFeign {

    @PostMapping("/errorType/findListByErrorTypeName")
    MSResponse<List<MDErrorType>> findListByErrorTypeName(@RequestParam("ids") List<Long> ids);

    @PostMapping("/errorCode/findListByErrorCodeName")
    MSResponse<List<MDErrorCode>> findListByErrorCodeName(@RequestParam("ids") List<Long> ids);

    @PostMapping("/actionCode/findListByActionCodeName")
    MSResponse<List<MDActionCode>> findListByActionCodeName(@RequestParam("ids") List<Long> ids);

}
