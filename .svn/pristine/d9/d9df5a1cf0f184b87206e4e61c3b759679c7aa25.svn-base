package com.kkl.kklplus.provider.rpt.chart.ms.md.service;

import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.feign.MSServicePointStreetFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

@Service
@Slf4j
public class MSServicePointStreetService {
    @Autowired
    private MSServicePointStreetFeign msServicePointStreetFeign;


    public List<NameValuePair<Integer, Long>> findAllServicePointQtyForRPT() {
        return MDUtils.findListUnnecessaryConvertType(() -> msServicePointStreetFeign.findServicePointStationCountListForRPT());
    }

    public Map<Long, List<NameValuePair<Integer, Long>>> findAllServicePointQtyCategoryForRPT() {
        return MDUtils.findMapServicePoint(() -> msServicePointStreetFeign.findStationCountListByProductCategoryForRPT());
    }

    public List<NameValuePair<Integer, Long>> findAllServicePointAutoPlanForRPT() {
        return MDUtils.findListUnnecessaryConvertType(() -> msServicePointStreetFeign.findStationCountListByAutoPlanFlagForRPT());
    }
}
