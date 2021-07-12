package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.md.MDServicePoint;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MDServicePointMapper extends CustomMapper<MDServicePoint, RPTServicePoint> {
    @Override
    public void mapAtoB(MDServicePoint a, RPTServicePoint b, MappingContext context) {
        b.setId(a.getId());
        b.setServicePointNo(a.getServicePointNo());
        b.setName(a.getName());
        b.setContactInfo1(a.getContactInfo1());
        b.setContactInfo2(a.getContactInfo2());
    }

    @Override
    public void mapBtoA(RPTServicePoint b, MDServicePoint a, MappingContext context) {
        a.setId(b.getId());
        a.setServicePointNo(b.getServicePointNo());
        a.setName(b.getName());
        a.setContactInfo1(b.getContactInfo1());
        a.setContactInfo2(b.getContactInfo2());
    }
}
