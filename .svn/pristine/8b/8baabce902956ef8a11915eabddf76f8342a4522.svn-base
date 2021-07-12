package com.kkl.kklplus.provider.rpt.chart.ms.md.service;

import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.provider.rpt.chart.ms.md.feign.MSServicePointQtyFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

@Service
@Slf4j
public class MSServicePointQtyService {

    @Autowired
    private MSServicePointQtyFeign msServicePointQtyFeign;


    public List<NameValuePair<Integer, Long>> findAllServicePointQtyForRPT() {
        return MDUtils.findListUnnecessaryConvertType(() -> msServicePointQtyFeign.findServicePointCountListByDegreeForRPT());
    }

    public Map<Long, List<NameValuePair<Integer, Integer>>> findAllServicePointQtyCategoryForRPT() {
        return MDUtils.findMapServicePoint(() -> msServicePointQtyFeign.findServicePointCountListByCategoryForRPT());
    }

    public List<NameValuePair<Long, Long>> findAllServicePointAutoPlanForRPT() {
        return MDUtils.findListUnnecessaryConvertType(() -> msServicePointQtyFeign.findServicePointCountListByCategoryAndAutoPlanFlagForRPT());
    }


}
