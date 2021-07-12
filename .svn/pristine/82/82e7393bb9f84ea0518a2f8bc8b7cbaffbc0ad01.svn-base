package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePoint;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.md.dto.MDServicePointAreaDto;
import com.kkl.kklplus.entity.md.dto.MDServicePointDto;
import com.kkl.kklplus.entity.md.dto.MDServicePointForRPTDto;
import com.kkl.kklplus.entity.rpt.web.RPTServicePointServiceArea;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSServicePointFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSServicePointFeignFallbackFactory.class)
public interface MSServicePointFeign {

    @PostMapping("/servicePoint/findBatchByIds")
    MSResponse<List<MDServicePoint>> findBatchByIds(@RequestBody List<Long> ids);

    /**
     * 返回网点数据
     * @param ids
     * fields
     * delFlag
     * @return
     * 网点id,Name,ServicePointNo
     */
    @PostMapping("/servicePoint/findBatchByIdsByCondition")
    MSResponse<List<MDServicePointViewModel>> findBatchByIdsByCondition(@RequestBody List<Long> ids, @RequestParam("fields") List<String> fields, @RequestParam("delFlag") Integer delFlag);

    /**
     * 分页获取网点信息
     * @param
     * @return
     */
    @PostMapping("/servicePointNew/findListForRPT")
    MSResponse<MSPage<MDServicePointForRPTDto>> findListForRPT(@RequestBody MSPage<MDServicePointForRPTDto> page, @RequestParam("areaIds") List<Long> areaIds, @RequestParam("servicePointId") Long servicePointId);


    /**
     * 分页获取区域ID列表复制
     * @param mdServicePointArea
     * @return
     */
    @PostMapping("/servicePointArea/findListWithAreas")
    MSResponse<MSPage<Long>> findListWithAreas(@RequestBody MDServicePointArea mdServicePointArea);

    /**
     * 查找网点覆盖的四级区域列表
     * @param pageNo
     * @param pageSize
     * @return
     */
    @GetMapping("/servicePointStation/findCoverAreaList")
    MSResponse<MSPage<Long>> findCoverAreaList(@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize);


    /**
     * 分页查找网点服务区域
     * @param
     * @return
     */
    @GetMapping("/servicePointAreaNew/findListByIdsAndAreaIdsForRPT")
    MSResponse<MSPage<MDServicePointArea>> findListByIdsAndAreaIdsForRPT(@RequestBody MDServicePointAreaDto servicePointAreaDto);


    /**
     * 分页查找网点服务区域
     * @param
     * @return
     */
    @GetMapping("servicePointArea/findAllListForRPT")
    MSResponse<MSPage<MDServicePointArea>> findServicePointArea(@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize);


    /**
     * 服务区域
     * @param
     * @return
     */
    @GetMapping("servicePointArea/findListByServicePointIdsForRPT")
    MSResponse<MSPage<MDServicePointArea>> findServicePointAreaByIds(@RequestBody List<Long> servicePointIds,@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize);
}
