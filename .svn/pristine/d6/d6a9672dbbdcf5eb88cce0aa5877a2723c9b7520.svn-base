package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSServiceTypeFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class MSServiceTypeService {

    @Autowired
    private MSServiceTypeFeign msServiceTypeFeign;

    /**
     * 获取所有数据
     *
     * @return
     */
    public List<RPTServiceType> findAllList() {
        return MDUtils.findAllList(RPTServiceType.class, msServiceTypeFeign::findAllList);
    }


}
