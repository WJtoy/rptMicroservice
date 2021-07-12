package com.kkl.kklplus.provider.rpt.ms.sys.service;

import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.entity.sys.SysUser;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSUserFeign;
import ma.glasnost.orika.MapperFacade;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author Zhoucy
 * @date 2018/9/25 10:18
 **/
@Service
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MSUserService {

    @Autowired
    private MSUserFeign userFeign;

    @Autowired
    private MapperFacade mapper;


    /**
     * 根据id列表查询用户信息
     */
    public Map<Long, RPTUser> getMapByUserIds(List<Long> userIds) {
        Map<Long, RPTUser> userMap = null;
        if (userIds != null && !userIds.isEmpty()) {
            MSResponse<Map<Long, SysUser>> responseEntity = userFeign.getMapByUserIds(userIds);
            if (MSResponse.isSuccess(responseEntity)) {
                userMap = responseEntity.getData().entrySet().stream().collect(Collectors.toMap(Map.Entry::getKey, i -> mapper.map(i.getValue(), RPTUser.class)));
            }
        }
        return userMap != null ? userMap : Maps.newHashMap();
    }

    /**
     * 根据id列表查询用户姓名
     */
    public Map<Long, String> getNamesByUserIds(List<Long> userIds) {
        Map<Long, String> nameMap = null;
        if (userIds != null && !userIds.isEmpty()) {
            MSResponse<Map<Long, String>> responseEntity = userFeign.getNamesByUserIds(userIds);
            if (MSResponse.isSuccess(responseEntity)) {
                nameMap = responseEntity.getData();
            }
        }
        return nameMap != null ? nameMap : Maps.newHashMap();
    }


    /**
     * 根据用户类型查询用户信息
     */
    public Map<Long, RPTUser> getMapByUserType(Integer userType) {
        Map<Long, RPTUser> userMap = null;
        if (userType != null) {
            MSResponse<Map<Long, SysUser>> responseEntity = userFeign.getMapByUserType(userType);
            if (MSResponse.isSuccess(responseEntity)) {
                userMap = responseEntity.getData().entrySet().stream().collect(Collectors.toMap(Map.Entry::getKey, i -> mapper.map(i.getValue(), RPTUser.class)));
            }
        }
        return userMap != null ? userMap : Maps.newHashMap();
    }

    /**
     * 根据用户类型查询用户名称
     */
    public Map<Long, String> getNameByUserType(Integer userType) {
        Map<Long, String> userMap = null;
        if (userType != null) {
            MSResponse<Map<Long, String>> responseEntity = userFeign.getNamesByUserType(userType);
            if (MSResponse.isSuccess(responseEntity)) {
                userMap = responseEntity.getData();
            }
        }
        return userMap != null ? userMap : Maps.newHashMap();
    }

}
