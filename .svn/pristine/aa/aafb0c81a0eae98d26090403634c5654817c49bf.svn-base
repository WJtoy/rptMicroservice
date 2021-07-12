package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSEngineerFeign;
import lombok.extern.slf4j.Slf4j;
import ma.glasnost.orika.MapperFacade;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@Service
@Slf4j
public class MSEngineerService {
    @Autowired
    private MSEngineerFeign msEngineerFeign;

    @Autowired
    private MapperFacade mapper;

    /**
     * servicePointService.getEngineersMap(engineerIds, Arrays.asList("id","name","appFlag"))
     */
    public Map<Long, RPTEngineer> getEngineersMap(List<Long> engineerIds, List<String> fields) {
        Map<Long, RPTEngineer> result = Maps.newHashMap();
        if (engineerIds != null && !engineerIds.isEmpty()) {
            List<RPTEngineer> list = findEngineersByIds(engineerIds, fields);
            if (list != null && !list.isEmpty()) {
                for (RPTEngineer item : list) {
                    result.put(item.getId(), item);
                }
            }
        }
        return result;
    }

    /**
     * servicePointService.findAllEngineersName(engineerIds, Arrays.asList("id","name","appFlag"))
     */
    public List<RPTEngineer> findAllEngineersName(List<Long> engineerIds, List<String> fields) {
        if (engineerIds != null && engineerIds.size()>0) {
            return findEngineersByIds(engineerIds, fields);
        }
        return Lists.newArrayList();
    }

    /**
     * 根据id列表获取安维人员列表
     * @param ids
     * @return
     */
    private List<RPTEngineer> findEngineersByIds(List<Long> ids, List<String> fields) {
        List<Field> fieldList = Lists.newArrayList();
        Class<?> cls = RPTEngineer.class;
        while(cls != null) {
            Field[] fields1 = cls.getDeclaredFields();
            fieldList.addAll(Arrays.asList(fields1));
            cls = cls.getSuperclass();
        }
        Long icount = fieldList.stream().filter(r->fields.contains(r.getName())).count();
        if (icount.intValue() != fields.size()) {
            throw new RuntimeException("按条件获取安维列表数据要求返回的字段有问题，请检查");
        }

        List<RPTEngineer> engineerList = Lists.newArrayList();
        if (ids != null && !ids.isEmpty()) {
            Lists.partition(ids,1000).forEach(longList -> {
                MSResponse<List<MDEngineer>> msResponse = msEngineerFeign.findEngineersByIds(longList, fields);
                if (MSResponse.isSuccess(msResponse)) {
                    List<RPTEngineer> engineerListFromMS = mapper.mapAsList(msResponse.getData(), RPTEngineer.class);
                    if (engineerListFromMS != null && !engineerListFromMS.isEmpty()) {
                        engineerList.addAll(engineerListFromMS);
                    }
                }
            });
        }
        return engineerList;
    }
}
