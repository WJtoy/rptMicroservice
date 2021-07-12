package com.kkl.kklplus.provider.rpt.service;

import com.kkl.kklplus.starter.redis.config.RedisGsonService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

/**
 * 报表服务基类
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class RptBaseService {

    protected static final int EXECL_CELL_WIDTH_15 = 15;
    protected static final int EXECL_CELL_WIDTH_10 = 10;
    protected static final int EXECL_CELL_WIDTH_20 = 20;
    protected static final int EXECL_CELL_HEIGHT_TITLE = 30;
    protected static final int EXECL_CELL_HEIGHT_HEADER = 20;
    protected static final int EXECL_CELL_HEIGHT_DATA = 20;

    @Autowired
    protected RedisGsonService redisGsonService;

}
