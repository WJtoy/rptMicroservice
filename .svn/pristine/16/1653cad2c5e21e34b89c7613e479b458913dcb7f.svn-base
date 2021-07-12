package com.kkl.kklplus.provider.rpt.utils;

import com.kkl.kklplus.entity.rpt.web.RPTUser;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.Map;
import java.util.Set;

@Service
@Slf4j
public class KeFuUtils {

    /**
     * 获取客服类型
     */
    public void getKeFu( Map<Long, RPTUser> kAKeFuMap,Integer subFlag, Set<Long> keFuIds) {

        for (RPTUser user : kAKeFuMap.values()) {
            if (user.getSubFlag() == subFlag) {
                keFuIds.add(user.getId());
            }
        }

    }
}
