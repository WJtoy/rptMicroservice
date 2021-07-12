package com.kkl.kklplus.provider.rpt.chart.ms.md.feign;


import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.fallback.MSServicePointStreetFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;

import java.util.List;
import java.util.Map;

@FeignClient(name = "provider-md", fallbackFactory = MSServicePointStreetFeignFallbackFactory.class)
public interface MSServicePointStreetFeign {


    /**
     * 获取常用网点街道，试用网点街道总数(20-常用，10-试用) ->图表
     */
    @GetMapping("/servicePointStation/findServicePointStationCountListForRPT")
    MSResponse<List<NameValuePair<Integer, Long>>> findServicePointStationCountListForRPT();


    /**
     * 按品类获取常用网点街道，试用网点街道总数 ->图表
     */
    @GetMapping("/servicePointStation/findStationCountListByProductCategoryForRPT")
    MSResponse<Map<Long, List<NameValuePair<Integer, Long>>>> findStationCountListByProductCategoryForRPT();

    /**
     * 获取自动派单街道数 1-自动派单街道总数，2-自动派单街道无常用网点，3-常用网点无自动派单   ->图表
     */
    @GetMapping("/servicePointStation/findStationCountListByAutoPlanFlagForRPT")
    MSResponse<List<NameValuePair<Integer, Long>>> findStationCountListByAutoPlanFlagForRPT();
}
