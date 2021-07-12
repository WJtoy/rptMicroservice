package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface ServicePointPraiseDetailsRptMapper {

    Page<RPTKeFuPraiseDetailsEntity> getServicePointPraiseDetailsPage(RPTKeFuCompleteTimeSearch search);

    Integer getServicePointPraiseSum(RPTKeFuCompleteTimeSearch search);

}
