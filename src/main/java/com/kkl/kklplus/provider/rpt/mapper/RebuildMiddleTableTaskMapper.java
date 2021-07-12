package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTRebuildMiddleTableTaskEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;


@Mapper
public interface RebuildMiddleTableTaskMapper {

    RPTRebuildMiddleTableTaskEntity get(@Param("id") Long id);


    Page<RPTRebuildMiddleTableTaskEntity> getTaskList(@Param("systemId") Integer systemId,
                                                      @Param("middleTableId") Integer middleTableId);

    void insert(RPTRebuildMiddleTableTaskEntity entity);

    void update(RPTRebuildMiddleTableTaskEntity entity);

}
