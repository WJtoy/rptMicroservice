package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.md.MDProduct;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MDProductMapper extends CustomMapper<MDProduct, RPTProduct> {
    @Override
    public void mapAtoB(MDProduct a, RPTProduct b, MappingContext context) {
        b.setId(a.getId());
        b.setName(a.getName());
    }

    @Override
    public void mapBtoA(RPTProduct b, MDProduct a, MappingContext context) {
        a.setId(b.getId());
        a.setName(b.getName());
    }
}
