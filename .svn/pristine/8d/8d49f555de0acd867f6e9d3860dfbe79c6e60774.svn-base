package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTReminderResponseTimeEntity;
import com.kkl.kklplus.entity.rpt.search.RPTReminderResponseTimeSearch;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface ReminderResponseTimeRptMapper {

    Page<RPTReminderResponseTimeEntity> getReminderResponseTimerListByPaging(RPTReminderResponseTimeSearch search);

    Integer hasReportData(RPTReminderResponseTimeSearch condition);
}
