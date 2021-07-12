package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.md.MDProductCategory;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSProductCategoryFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSProductCategoryFeignFallbackFactory implements FallbackFactory<MSProductCategoryFeign> {
    @Override
    public MSProductCategoryFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSProductCategoryFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSProductCategoryFeign() {
            @Override
            public MSResponse<List<MDProductCategory>> findAllList() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }



            @Override
            public MSResponse<Long> getIdByCode(String code) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Long> getIdByName(String code) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Integer> delete(MDProductCategory mdProductCategory) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDProductCategory>> findList(MDProductCategory mdProductCategory) {
                log.warn("{}", throwable);
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<NameValuePair<Long, String>>> findAllListForRPT() {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
            @Override
            public MSResponse<MDProductCategory> getById(Long id) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Integer> insert(MDProductCategory mdProductCategory) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Integer> update(MDProductCategory mdProductCategory) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
