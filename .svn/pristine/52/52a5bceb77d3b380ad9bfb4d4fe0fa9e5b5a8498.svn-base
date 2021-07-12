package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTComplainStatisticsDailyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTComplainStatisticsDailySearch;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface ComplainStatisticsDailyMapper {

    List<RPTComplainStatisticsDailyEntity> getDayComplainSumNew(RPTComplainStatisticsDailySearch search);


    Integer hasReportData(RPTComplainStatisticsDailySearch search);
}
