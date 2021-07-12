package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.md.MDCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;
import org.springframework.stereotype.Component;

@Component
public class MDCustomerMapper extends CustomMapper<MDCustomer, RPTCustomer> {
    @Override
    public void mapAtoB(MDCustomer a, RPTCustomer b, MappingContext context) {
        b.setId(a.getId());
        b.setCode(a.getCode());
        b.setName(a.getName());
        b.setFullName(a.getFullName());
        b.setSales(a.getSalesId() == null ? null : new RPTUser(a.getSalesId()));
        b.setMerchandiser(a.getMerchandiserId() == null ? null: new RPTUser(a.getMerchandiserId()));
        b.setContractDate(a.getContractDate());
        int paymentType = a.getPaymentType() == null ? 0 : a.getPaymentType();
        b.setPaymentType(new RPTDict(paymentType, ""));
        b.setRemarks(a.getRemarks());
    }

    @Override
    public void mapBtoA(RPTCustomer b, MDCustomer a, MappingContext context) {
        a.setId(b.getId());
        a.setCode(b.getCode());
        a.setName(b.getName());
        a.setFullName(b.getFullName());
        a.setSalesId(b.getSales() == null ? null : b.getSales().getId());
        a.setMerchandiserId(b.getMerchandiser() == null ? null: b.getMerchandiser().getId());
        a.setContractDate(b.getContractDate());
        int paymentType = b.getPaymentType() != null? StringUtils.toInteger(b.getPaymentType().getValue()) : 0;
        a.setPaymentType(paymentType);
        a.setRemarks(b.getRemarks());
    }
}
