package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDMaterial;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSMaterialMasterFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;
@Component
@Slf4j
public class MSMaterialMasterFallbackFactory implements FallbackFactory<MSMaterialMasterFeign> {
    @Override
    public MSMaterialMasterFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSMaterialMasterFallbackFactory:{}",throwable.getMessage());
        }

        return new MSMaterialMasterFeign() {
            @Override
            public MSResponse<List<MDMaterial>> findMaterialById(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
