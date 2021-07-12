package com.kkl.kklplus.provider.rpt.ms.sys.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.sys.SysUser;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSUserFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

/**
 * @author Zhoucy
 * @date 2018/9/25 10:12
 **/
@Component
@Slf4j
public class MSUserFeignFallbackFactory implements FallbackFactory<MSUserFeign> {

    @Override
    public MSUserFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSUserFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSUserFeign() {
            @Override
            public MSResponse<Map<Long, SysUser>> getMapByUserType(Integer userType) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Map<Long, String>> getNamesByUserType(Integer userType) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Map<Long, SysUser>> getMapByUserIds(List<Long> userIds) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<Map<Long, String>> getNamesByUserIds(List<Long> userIds) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
