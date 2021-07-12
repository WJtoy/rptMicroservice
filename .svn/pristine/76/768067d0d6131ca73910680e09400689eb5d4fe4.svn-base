package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDMaterial;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSMaterialMasterFeign;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

@Service
@Slf4j
public class MSMaterialMasterService {

    @Autowired
    private MSMaterialMasterFeign msMaterialMasterFeign;

    public List<MDMaterial> findMaterialWithIds(List<Long> ids) {
        List<MDMaterial> MDMaterialList = Lists.newArrayList();
        if (ids != null && !ids.isEmpty()) {
            ids = ids.stream().distinct().collect(Collectors.toList());
            if (!ids.isEmpty()) {
                List<List<Long>> materialIds = Lists.partition(ids, 10);
                materialIds.forEach(longList -> {
                    MSResponse<List<MDMaterial>> mdMaterials = msMaterialMasterFeign.findMaterialById(longList);
                    Optional.ofNullable(mdMaterials.getData()).ifPresent(MDMaterialList::addAll);
                });
            }
        }
        return MDMaterialList;
    }

}
