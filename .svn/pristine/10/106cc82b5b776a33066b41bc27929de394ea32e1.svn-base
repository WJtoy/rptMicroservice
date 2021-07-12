package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.UserCustomerEntity;
import com.kkl.kklplus.provider.rpt.mapper.SysUserCustomerMappingMapper;
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
public class SysUserCustomerMappingService extends RptBaseService {

    @Resource
    private SysUserCustomerMappingMapper sysUserCustomerMappingMapper;

    /**
     * 从Web数据库提取用户客户关系
     */
    private List<UserCustomerEntity> getSysUserCustomerMappingListFromWebDB() {
        List<UserCustomerEntity> result = Lists.newArrayList();
        List<UserCustomerEntity> list = sysUserCustomerMappingMapper.getSysUserCustomerMappingFromWebDB();
        if (list != null && !list.isEmpty()) {
            result.addAll(list);
        }
        return result;
    }

    private void insertSysUserCustomerMapping(UserCustomerEntity entity, int systemId, List<String> quarters) {
        try {

            for (String quarter : quarters) {
                entity.setSystemId(systemId);
                entity.setQuarter(quarter);
                sysUserCustomerMappingMapper.insert(entity);
            }
        } catch (Exception e) {
            log.error("【SysUserCustomerMappingService.insertSysUserCustomerMapping】OrderId: {}, errorMsg: {}", entity.getUserId(), Exceptions.getStackTraceAsString(e));
        }

    }


    /**
     *
     */
    public void saveSysUserCustomerMappingsToRptDB() {
        List<UserCustomerEntity> list = getSysUserCustomerMappingListFromWebDB();
        if (!list.isEmpty()) {
            List<String> quarters = QuarterUtils.getRptQuarters();
            int systemId = RptCommonUtils.getSystemId();
            for (UserCustomerEntity item : list) {
                insertSysUserCustomerMapping(item, systemId, quarters);
            }
        }

    }

    /**
     *
     */
    public void deleteSysUserCustomerMappingsFromRptDB() {
        List<String> quarters = QuarterUtils.getRptQuarters();
        int systemId = RptCommonUtils.getSystemId();
        for (String quarter : quarters) {
            sysUserCustomerMappingMapper.deleteSysUserCustomerMappings(quarter, systemId);
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
                        saveSysUserCustomerMappingsToRptDB();
                        break;
                    case INSERT_MISSED_DATA:
                        break;
                    case UPDATE:
                        deleteSysUserCustomerMappingsFromRptDB();
                        saveSysUserCustomerMappingsToRptDB();
                        break;
                    case DELETE:
                        deleteSysUserCustomerMappingsFromRptDB();
                        break;
                }
                result = true;
            } catch (Exception e) {
                log.error("SysUserCustomerMappingService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
}
