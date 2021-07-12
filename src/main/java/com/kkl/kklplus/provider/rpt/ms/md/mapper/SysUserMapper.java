package com.kkl.kklplus.provider.rpt.ms.md.mapper;

import com.kkl.kklplus.entity.rpt.web.RPTUser;
import com.kkl.kklplus.entity.sys.SysUser;
import ma.glasnost.orika.CustomMapper;
import ma.glasnost.orika.MappingContext;

/**
 * @author Zhoucy
 * @date 2018/9/25 10:21
 **/
public class SysUserMapper extends CustomMapper<SysUser, RPTUser> {

    @Override
    public void mapAtoB(SysUser a, RPTUser b, MappingContext context) {
        b.setId(a.getId());
        b.setName(a.getName());
        b.setEmail(a.getEmail());
        b.setPhone(a.getPhone());
        b.setMobile(a.getMobile());
        b.setQq(a.getQq());
    }

    @Override
    public void mapBtoA(RPTUser b, SysUser a, MappingContext context) {
        a.setId(b.getId());
        a.setName(b.getName());
        a.setEmail(b.getEmail());
        a.setPhone(b.getPhone());
        a.setMobile(b.getMobile());
        a.setQq(b.getQq());
    }
}
