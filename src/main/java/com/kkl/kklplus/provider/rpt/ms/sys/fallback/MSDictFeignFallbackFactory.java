package com.kkl.kklplus.provider.rpt.ms.sys.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.sys.SysDict;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSDictFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSDictFeignFallbackFactory implements FallbackFactory<MSDictFeign> {

    private static String errorMsg = "操作超时";

    @Override
    public MSDictFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSDictFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSDictFeign() {
            @Override
            public MSResponse<SysDict> get(Long id) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<List<SysDict>> getAllList() {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<MSPage<SysDict>> getList(SysDict dict) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<List<SysDict>> getListByType(String type) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<List<String>> getTypeList() {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<SysDict> insert(SysDict dict) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<Integer> update(SysDict dict) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<Integer> delete(SysDict dict) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<Boolean> reloadAllToRedis() {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }

            @Override
            public MSResponse<Boolean> reloadToRedis(String type) {
                return new MSResponse<>(MSErrorCode.newInstance(MSErrorCode.FAILURE, errorMsg));
            }
        };
    }
}
