package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.md.MDProductCategory;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MDProductCategoryMapper extends CustomMapper<MDProductCategory, RPTProductCategory> {
    @Override
    public void mapAtoB(MDProductCategory a, RPTProductCategory b, MappingContext context) {
        b.setId(a.getId());
        b.setName(a.getName());
        b.setCode(a.getCode());
    }

    @Override
    public void mapBtoA(RPTProductCategory b, MDProductCategory a, MappingContext context) {
        a.setId(b.getId());
        a.setName(b.getName());
        a.setCode(b.getCode());
    }
}
