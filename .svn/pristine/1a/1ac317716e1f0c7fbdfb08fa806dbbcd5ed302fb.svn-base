package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTExportTaskEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;


@Mapper
public interface RptExportTaskMapper {

    RPTExportTaskEntity get(@Param("id") Long id);

    List<RPTExportTaskEntity> getTaskListByHashCode(@Param("systemId") Integer systemId,
                                                    @Param("reportId") Integer reportId,
                                                    @Param("taskCreateDate") Long taskCreateDate,
                                                    @Param("taskCreateBy") Long taskCreateBy,
                                                    @Param("searchConditionHashcode") Integer searchConditionHashcode);

    Page<RPTExportTaskEntity> getTaskList(@Param("systemId") Integer systemId,
                                          @Param("taskCreateBy") Long taskCreateBy,
                                          @Param("reportId") Integer reportId,
                                          @Param("reportType") Integer reportType,
                                          @Param("beginTaskCreateDate") Long beginTaskCreateDate,
                                          @Param("endTaskCreateDate") Long endTaskCreateDate);

    void insert(RPTExportTaskEntity entity);

    void update(RPTExportTaskEntity entity);

    void updateDownloadInfo(RPTExportTaskEntity entity);

}
