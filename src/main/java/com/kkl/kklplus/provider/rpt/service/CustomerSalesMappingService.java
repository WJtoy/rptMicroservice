package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTCustomerSalesMappingEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.provider.rpt.mapper.CustomerSalesMappingMapper;
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
public class CustomerSalesMappingService extends RptBaseService {

    @Resource
    private CustomerSalesMappingMapper customerSalesMappingMapper;

    /**
     * 从Web数据库提取客户业务员映射关系
     */
    private List<RPTCustomerSalesMappingEntity> getCustomerSalesMappingListFromWebDB() {
        List<RPTCustomerSalesMappingEntity> result = Lists.newArrayList();
        List<RPTCustomerSalesMappingEntity> list = customerSalesMappingMapper.getCustomerSalesMappingFromWebDB();
        if (list != null && !list.isEmpty()) {
            result.addAll(list);
        }
        return result;
    }

    private void insertCustomerSalesMapping(RPTCustomerSalesMappingEntity entity, int systemId, List<String> quarters) {
        try {

            for (String quarter : quarters) {
                entity.setSystemId(systemId);
                entity.setQuarter(quarter);
                customerSalesMappingMapper.insert(entity);
            }
        } catch (Exception e) {
            log.error("【CustomerSalesMappingService.insertCustomerSalesMapping】OrderId: {}, errorMsg: {}", entity.getCustomerId(), Exceptions.getStackTraceAsString(e));
        }

    }


    /**
     *
     */
    public void saveCustomerSalesMappingsToRptDB() {
        List<RPTCustomerSalesMappingEntity> list = getCustomerSalesMappingListFromWebDB();
        if (!list.isEmpty()) {
            List<String> quarters = QuarterUtils.getRptQuarters();
            int systemId = RptCommonUtils.getSystemId();
            for (RPTCustomerSalesMappingEntity item : list) {
                insertCustomerSalesMapping(item, systemId, quarters);
            }
        }

    }

    /**
     *
     */
    public void deleteCustomerSalesMappingsFromRptDB() {
        List<String> quarters = QuarterUtils.getRptQuarters();
        int systemId = RptCommonUtils.getSystemId();
        for (String quarter : quarters) {
            customerSalesMappingMapper.deleteCustomerSalesMappings(quarter, systemId);
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
                        saveCustomerSalesMappingsToRptDB();
                        break;
                    case INSERT_MISSED_DATA:
                        break;
                    case UPDATE:
                        deleteCustomerSalesMappingsFromRptDB();
                        saveCustomerSalesMappingsToRptDB();
                        break;
                    case DELETE:
                        deleteCustomerSalesMappingsFromRptDB();
                        break;
                }
                result = true;
            } catch (Exception e) {
                log.error("CustomerSalesMappingService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
}
