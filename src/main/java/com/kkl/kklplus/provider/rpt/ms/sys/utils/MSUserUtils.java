package com.kkl.kklplus.provider.rpt.ms.sys.utils;

import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSUserService;
import com.kkl.kklplus.provider.rpt.utils.SpringContextHolder;

import java.util.List;
import java.util.Map;

/**
 * @author Zhoucy
 * @date 2018/9/25 11:17
 **/
public class MSUserUtils {

    private static MSUserService msUserService = SpringContextHolder.getBean(MSUserService.class);


    public static Map<Long, RPTUser> getMapByUserType(Integer userType) {
        return msUserService.getMapByUserType(userType);
    }

    public static Map<Long, String> getNamesByUserType(Integer userType) {
        return msUserService.getNameByUserType(userType);
    }

    public static Map<Long, RPTUser> getMapByUserIds(List<Long> userIds) {
        return msUserService.getMapByUserIds(userIds);
    }

    public static Map<Long, String> getNamesByUserIds(List<Long> userIds) {
        return msUserService.getNamesByUserIds(userIds);
    }

}
