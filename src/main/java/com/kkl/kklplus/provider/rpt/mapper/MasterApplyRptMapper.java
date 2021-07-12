package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTKeFuPraiseDetailsEntity;
import com.kkl.kklplus.entity.rpt.RPTMasterApplyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTKeFuCompleteTimeSearch;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface MasterApplyRptMapper {

    Page<RPTMasterApplyEntity> getMasterApplyPage(@Param("beginDate") Date beginDate,
                                                  @Param("endDate") Date endDate,
                                                  @Param("customerId") Long customerId,
                                                  @Param("quarter") String quarter);



    List<RPTMasterApplyEntity> getMasterApplyList(@Param("beginDate") Date beginDate,
                                                  @Param("endDate") Date endDate,
                                                  @Param("customerId") Long customerId,
                                                  @Param("quarter") String quarter);



    Integer hasReportData(@Param("beginDate") Date beginDate,
                          @Param("endDate") Date endDate,
                          @Param("customerId") Long customerId,
                          @Param("quarter") String quarter);


}
