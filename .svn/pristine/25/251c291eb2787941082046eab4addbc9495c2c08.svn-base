package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.md.MDEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MDEngineerMapper extends CustomMapper<MDEngineer, RPTEngineer> {
    @Override
    public void mapAtoB(MDEngineer a, RPTEngineer b, MappingContext context) {
        b.setId(a.getId());
        b.setName(a.getName());
        b.setMasterFlag(a.getMasterFlag());
        b.setAppFlag(a.getAppFlag());
    }

    @Override
    public void mapBtoA(RPTEngineer b, MDEngineer a, MappingContext context) {
        a.setId(b.getId());
        a.setName(b.getName());
        a.setMasterFlag(b.getMasterFlag());
        a.setAppFlag(b.getAppFlag());
    }
}
