package com.kkl.kklplus.provider.rpt.ms.sys.mapper;

import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.sys.SysDict;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MSDictMapper extends CustomMapper<SysDict, RPTDict> {

    @Override
    public void mapAtoB(SysDict a, RPTDict b, MappingContext context) {
        b.setId(a.getId());
        b.setLabel(a.getLabel());
        b.setValue(a.getValue());
        b.setType(a.getType());
    }

    @Override
    public void mapBtoA(RPTDict b, SysDict a, MappingContext context) {
        a.setId(b.getId());
        a.setLabel(b.getLabel());
        a.setValue(b.getValue());
        a.setType(b.getType());
    }
}
