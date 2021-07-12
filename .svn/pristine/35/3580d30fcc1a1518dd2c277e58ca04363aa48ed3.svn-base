package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface KeFuPraiseDetailsRptMapper {

    Page<RPTKeFuPraiseDetailsEntity> getKeFuPraiseDetailsPage(RPTKeFuCompleteTimeSearch search);

    Integer getKeFuPraiseSum(RPTKeFuCompleteTimeSearch search);

}
