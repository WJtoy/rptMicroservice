package com.kkl.kklplus.provider.rpt.common.service;

import com.kkl.kklplus.provider.rpt.common.mapper.SysUserSubMapper;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.List;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class SysUserSubService {


    @Resource
    private SysUserSubMapper mapper;

    /**
     * 读取账号负责的子账号列表
     * @param userId
     * @return
     */
    public List<Long> getSubUserIdsOfManager(Long userId){
        return mapper.getSubUserIdsOfManager(userId);
    }

}
