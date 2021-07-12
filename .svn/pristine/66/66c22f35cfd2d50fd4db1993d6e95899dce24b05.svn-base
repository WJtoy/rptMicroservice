package com.kkl.kklplus.provider.rpt.ms.sys.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePointStation;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.sys.SysArea;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSAreaFeign;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSDictFeign;
import com.kkl.kklplus.starter.redis.config.RedisGsonService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.redis.core.RedisCallback;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Set;

/**
 * 区域Service
 */
@Slf4j
@Service
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MSAreaService {

    public static final int REDIS_SYS_DB = 1;
    public static final String SYS_AREA_TYPE = "area:type:%s";

    @Autowired
    public RedisTemplate redisTemplate;

    @Autowired
    public MSAreaFeign msAreaFeign;


    @Autowired
    public MSDictFeign msDictFeign;


    @Autowired
    private RedisGsonService redisGsonService;


    /**
     * 按区域类型返回所有区域清单
     */
    public List<RPTArea> getAreaByType(Integer type) {
        List<RPTArea> result = Lists.newArrayList();
        if (type != null && type >=  RPTArea.TYPE_VALUE_PROVINCE && type <= RPTArea.TYPE_VALUE_TOWN) {
            byte[] bKey = String.format(SYS_AREA_TYPE, type).getBytes(StandardCharsets.UTF_8);
            Set set = null;
            try {
                set = (Set) redisTemplate.execute((RedisCallback<Object>) connection -> {
                    connection.select(REDIS_SYS_DB);
                    return connection.zRange(bKey, 0, -1);
                });
            } catch (Exception e) {
                log.error("[RedisUtils.getAreaByType]", e);
            }
            if (set != null && !set.isEmpty()) {
                RPTArea area;
                for (Object item : set) {
                    if (item instanceof byte[]) {
                        area = redisGsonService.fromJson((byte[]) item, RPTArea.class);
                        if (area != null) {
                            result.add(area);
                        }
                    }
                }
            }
        }
        return result;
    }

    /**
     * 查找突击覆盖的四级区域列表
     *
     * @return
     */
    public List<Long> findCoverAreaList(List<Long> productCategoryIds) {
        List<Long> subAreaIdList = Lists.newArrayList();
        int pageNo = 1;
        int pageSize = 2000;

        MSResponse<MSPage<Long>> msResponse = msAreaFeign.getCrushList(pageNo, pageSize,productCategoryIds);
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<Long> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null) {
                subAreaIdList.addAll(returnPage.getList());

                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    MSResponse<MSPage<Long>> whileMSResponse = msAreaFeign.getCrushList(pageNo, pageSize,productCategoryIds);
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<Long> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null) {
                            subAreaIdList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }
        }
        return subAreaIdList;
    }

    /**
     * 查找远程覆盖的四级区域列表
     *
     * @return
     */
    public List<Long> findTravelAreaList(List<Long> productCategoryIds) {
        List<Long> subAreaIdList = Lists.newArrayList();
        int pageNo = 1;
        int pageSize = 2000;

        MSResponse<MSPage<Long>> msResponse = msAreaFeign.getTravelList(pageNo, pageSize,productCategoryIds);
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<Long> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null) {
                subAreaIdList.addAll(returnPage.getList());

                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    MSResponse<MSPage<Long>> whileMSResponse = msAreaFeign.getTravelList(pageNo, pageSize,productCategoryIds);
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<Long> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null) {
                            subAreaIdList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }
        }
        return subAreaIdList;
    }

}
