package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTExploitDetailEntity;
import com.kkl.kklplus.entity.rpt.search.RPTExploitDetailSearch;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface ExploitDetailRptMapper {

    Page<RPTExploitDetailEntity> getExploitDetailList(RPTExploitDetailSearch search);

    Integer hasReportData(RPTExploitDetailSearch search);
    
}
