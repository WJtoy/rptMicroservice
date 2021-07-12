package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.CustomerInfoEntity;
import com.kkl.kklplus.provider.rpt.mapper.CustomerInfoMappingMapper;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.List;


@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerInfoMappingService extends RptBaseService {

    @Resource
    private CustomerInfoMappingMapper customerInfoMappingMapper;

    /**
     * 从Web数据库提取VIp客户
     */
    private List<CustomerInfoEntity> getCustomerInfoMappingListFromWebDB() {
        List<CustomerInfoEntity> result = Lists.newArrayList();
        List<CustomerInfoEntity> list = customerInfoMappingMapper.getCustomerInfoMappingFromWebDB();
        if (list != null && !list.isEmpty()) {
            result.addAll(list);
        }
        return result;
    }

    private void insertCustomerInfoMapping(CustomerInfoEntity entity, int systemId, List<String> quarters) {
        try {

            for (String quarter : quarters) {
                entity.setSystemId(systemId);
                entity.setQuarter(quarter);
                customerInfoMappingMapper.insert(entity);
            }
        } catch (Exception e) {
            log.error("【CustomerInfoMappingService.insertCustomerInfoMapping】OrderId: {}, errorMsg: {}", entity.getCustomerId(), Exceptions.getStackTraceAsString(e));
        }

    }


    /**
     *
     */
    public void saveCustomerInfoMappingsToRptDB() {
        List<CustomerInfoEntity> list = getCustomerInfoMappingListFromWebDB();
        if (!list.isEmpty()) {
            List<String> quarters = QuarterUtils.getRptQuarters();
            int systemId = RptCommonUtils.getSystemId();
            for (CustomerInfoEntity item : list) {
                insertCustomerInfoMapping(item, systemId, quarters);
            }
        }

    }

    /**
     *
     */
    public void deleteCustomerInfoMappingsFromRptDB() {
        List<String> quarters = QuarterUtils.getRptQuarters();
        int systemId = RptCommonUtils.getSystemId();
        for (String quarter : quarters) {
            customerInfoMappingMapper.deleteCustomerInfoMappings(quarter, systemId);
        }
    }

    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Long beginDt, Long endDt) {
        boolean result = false;
        if (operationType != null && beginDt != null && beginDt > 0 && endDt != null && endDt > 0) {
            try {

                switch (operationType) {
                    case INSERT:
                        saveCustomerInfoMappingsToRptDB();
                        break;
                    case INSERT_MISSED_DATA:
                        break;
                    case UPDATE:
                        deleteCustomerInfoMappingsFromRptDB();
                        saveCustomerInfoMappingsToRptDB();
                        break;
                    case DELETE:
                        deleteCustomerInfoMappingsFromRptDB();
                        break;
                }
                result = true;
            } catch (Exception e) {
                log.error("CustomerInfoMappingService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
}
