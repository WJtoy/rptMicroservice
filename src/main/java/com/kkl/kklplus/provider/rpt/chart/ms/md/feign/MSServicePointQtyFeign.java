package com.kkl.kklplus.provider.rpt.chart.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.fallback.MSServicePointQtyFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;

import java.util.List;
import java.util.Map;

@FeignClient(name = "provider-md", fallbackFactory = MSServicePointQtyFeignFallbackFactory.class)
public interface MSServicePointQtyFeign {


    /**
     * 获取网点数量根据网点分级 ->图表
     */
    @GetMapping("/servicePointNew/findServicePointCountListByDegreeForRPT")
    MSResponse<List<NameValuePair<Integer,Long>>> findServicePointCountListByDegreeForRPT();

    /**
     * 查询网点数量根据品类 ->图表
     */
    @GetMapping("/servicePointNew/findServicePointCountListByCategoryForRPT")
    MSResponse<Map<Long, List<NameValuePair<Integer, Integer>>>> findServicePointCountListByCategoryForRPT();

    /**
     * 查询网点数量根据品类以及自动派单 ->图表
     */
    @GetMapping("/servicePointNew/findServicePointCountListByCategoryAndAutoPlanFlagForRPT")
    MSResponse<List<NameValuePair<Long, Long>>> findServicePointCountListByCategoryAndAutoPlanFlagForRPT();

}
