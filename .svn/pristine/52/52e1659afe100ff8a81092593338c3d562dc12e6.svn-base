package com.kkl.kklplus.provider.rpt.ms.sys.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.sys.SysUser;
import com.kkl.kklplus.provider.rpt.ms.sys.fallback.MSUserFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;

import java.util.List;
import java.util.Map;

/**
 * @author Zhoucy
 * @date 2018/9/25 10:11
 **/
@FeignClient(name = "provider-user", fallbackFactory = MSUserFeignFallbackFactory.class)
public interface MSUserFeign {

    /**
     * 根据用户类型获取用户信息
     */
    @GetMapping("/user/getMapByUserType/{userType}")
    MSResponse<Map<Long, SysUser>> getMapByUserType(@PathVariable("userType") Integer userType);

    /**
     * 根据用户类型获取用户名称
     */
    @GetMapping("/user/getNamesByUserType/{userType}")
    MSResponse<Map<Long, String>> getNamesByUserType(@PathVariable("userType") Integer userType);

    /**
     * 根据用户Id列表获取用户信息
     */
    @PostMapping("/user/getMapByUserIds")
    MSResponse<Map<Long, SysUser>> getMapByUserIds(@RequestBody List<Long> userIds);

    /**
     * 根据用户Id列表获取用户名称
     */
    @PostMapping("/user/getNamesByUserIds")
    MSResponse<Map<Long, String>> getNamesByUserIds(@RequestBody List<Long> userIds);

}
