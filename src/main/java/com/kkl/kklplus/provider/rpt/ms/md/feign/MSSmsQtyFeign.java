package com.kkl.kklplus.provider.rpt.ms.md.feign;


import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSSmsQtyFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.Map;


@FeignClient(name = "provider-sms", fallbackFactory = MSSmsQtyFeignFallbackFactory.class)
public interface MSSmsQtyFeign {

    @PostMapping("/smsQty/shortMessageCache")
    MSResponse<Map<Integer, Long>> shortMessageCache(@RequestParam("date") String date);
}
