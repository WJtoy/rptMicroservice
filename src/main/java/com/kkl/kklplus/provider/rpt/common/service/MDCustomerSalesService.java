package com.kkl.kklplus.provider.rpt.common.service;

import com.kkl.kklplus.provider.rpt.common.mapper.MDCustomerSalesMapper;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.List;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MDCustomerSalesService {


    @Resource
    private MDCustomerSalesMapper mapper;

    /**
     * 读取所有业务Id列表
     * @param systemId
     * @return
     */
    public List<Long> getSalesIdList(Integer systemId){
        List<Long> ids = mapper.getSalesIdList(systemId);
        return ids.stream().distinct().collect(Collectors.toList());
    }

}
