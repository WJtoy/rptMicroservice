package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePoint;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.md.dto.MDServicePointAreaDto;
import com.kkl.kklplus.entity.md.dto.MDServicePointForRPTDto;
import com.kkl.kklplus.entity.rpt.web.RPTServicePointServiceArea;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSServicePointFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSServicePointFeignFallbackFactory implements FallbackFactory<MSServicePointFeign> {
    @Override
    public MSServicePointFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSServicePointFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSServicePointFeign() {
            @Override
            public MSResponse<List<MDServicePoint>> findBatchByIds(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<MDServicePointViewModel>> findBatchByIdsByCondition(List<Long> ids, List<String> fields, Integer delFlag) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDServicePointForRPTDto>> findListForRPT(MSPage<MDServicePointForRPTDto> page, List<Long> areaIds, Long servicePointId) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<Long>> findListWithAreas(MDServicePointArea mdServicePointArea) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<Long>> findCoverAreaList(int pageNo, int pageSize) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDServicePointArea>> findListByIdsAndAreaIdsForRPT(MDServicePointAreaDto servicePointAreaDto) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDServicePointArea>> findServicePointArea(int pageNo, int pageSize) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDServicePointArea>> findServicePointAreaByIds(List<Long> servicePointIds, int pageNo, int pageSize) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }


        };
    }
}
