package com.kkl.kklplus.provider.rpt.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.provider.rpt.entity.SysUserRegionEntity;
import com.kkl.kklplus.provider.rpt.mapper.SysUserRegionMappingMapper;
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
public class SysUserRegionMappingService extends RptBaseService {

    @Resource
    private SysUserRegionMappingMapper sysUserRegionMappingMapper;

    /**
     * 从Web数据库提取客户业务员映射关系
     */
    private List<SysUserRegionEntity> getSysUserRegionMappingListFromWebDB() {
        List<SysUserRegionEntity> result = Lists.newArrayList();
        List<SysUserRegionEntity> list = sysUserRegionMappingMapper.getSysUserRegionMappingFromWebDB();
        if (list != null && !list.isEmpty()) {
            result.addAll(list);
        }
        return result;
    }

    private void insertSysUserRegionMapping(SysUserRegionEntity entity, int systemId, List<String> quarters) {
        try {

            for (String quarter : quarters) {
                entity.setSystemId(systemId);
                entity.setQuarter(quarter);
                sysUserRegionMappingMapper.insert(entity);
            }
        } catch (Exception e) {
            log.error("【SysUserRegionMappingService.insertSysUserRegionMapping】OrderId: {}, errorMsg: {}", entity.getUserId(), Exceptions.getStackTraceAsString(e));
        }

    }


    /**
     *
     */
    public void saveSysUserRegionMappingsToRptDB() {
        List<SysUserRegionEntity> list = getSysUserRegionMappingListFromWebDB();
        if (!list.isEmpty()) {
            List<String> quarters = QuarterUtils.getRptQuarters();
            int systemId = RptCommonUtils.getSystemId();
            for (SysUserRegionEntity item : list) {
                insertSysUserRegionMapping(item, systemId, quarters);
            }
        }

    }

    /**
     *
     */
    public void deleteSysUserRegionMappingsFromRptDB() {
        List<String> quarters = QuarterUtils.getRptQuarters();
        int systemId = RptCommonUtils.getSystemId();
        for (String quarter : quarters) {
            sysUserRegionMappingMapper.deleteSysUserRegionMappings(quarter, systemId);
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
                        saveSysUserRegionMappingsToRptDB();
                        break;
                    case INSERT_MISSED_DATA:
                        break;
                    case UPDATE:
                        deleteSysUserRegionMappingsFromRptDB();
                        saveSysUserRegionMappingsToRptDB();
                        break;
                    case DELETE:
                        deleteSysUserRegionMappingsFromRptDB();
                        break;
                }
                result = true;
            } catch (Exception e) {
                log.error("SysUserRegionMappingService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }
}
